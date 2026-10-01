"""Extracción de datos de boletas con un modelo de visión local (Ollama).

Pensado para correr en CPU con ~8 GB de RAM:
  - modelo liviano por defecto (gemma3:4b)
  - se recorta la boleta de la foto (descartando boletas vecinas y el fondo) y se procesa
    en dos franjas (encabezado y pie) para que el modelo vea el texto chico con más resolución
  - validaciones cruzadas (IVA vs total, fecha reciente, RUT) que marcan en la columna
    'revisar' lo que conviene mirar a mano
  - un solo modelo cargado, con keep_alive para no recargarlo entre boletas

Uso:
    uv add ollama pandas pydantic openpyxl pillow
    ollama pull gemma3:4b
    uv run main.py
    uv run main.py --carpeta boletas --salida boletas_extraidas.xlsx --debug
"""

from __future__ import annotations

import argparse
import io
import re
import sys
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import ollama
import pandas as pd
from PIL import Image, ImageOps
from pydantic import BaseModel, ConfigDict, Field, ValidationError, field_validator, model_validator

MODELO_POR_DEFECTO = "gemma3:4b"
EXTENSIONES_VALIDAS = {".jpg", ".jpeg", ".png", ".webp"}
MAX_LADO_PX = 1280

# temperature=0 (greedy) hace que los modelos chicos caigan en bucles de repetición.
OPCIONES_GENERACION = {
    "temperature": 0.1,
    "repeat_penalty": 1.1,
    "num_predict": 512,  # el JSON cabe de sobra; corta bucles
    "num_ctx": 4096,  # contexto acotado: menos RAM y más rápido en CPU
}

PISTAS = {
    "nombre_emisor": "razón social completa del comercio (nombre grande bajo el logo o bajo 'BOLETA ELECTRÓNICA'); "
    "puede ocupar dos líneas, cópialo entero. No es el giro ni la dirección.",
    "rut_emisor": "RUT del emisor, con formato como 77.648.040-1. Puede venir junto a 'R.U.T.' o "
    "justo bajo el nombre del comercio, sin etiqueta.",
    "fecha": "fecha de emisión (línea 'Fecha' o 'Fecha de emisión'), convertida a YYYY-MM-DD, sin la hora.",
    "folio_boleta": "número de la boleta, escrito después de 'BOLETA ELECTRÓNICA' o de 'N°'. "
    "No es la caja, la mesa, el pedido ni el POS. Solo dígitos.",
    "monto_total": "valor de 'Total' o 'Monto Total', no el neto ni el IVA ni un ítem suelto.",
    "iva": "valor del IVA (línea 'IVA' o 'El IVA de esta boleta es'), no el neto ni el total.",
    "linea_pago": "copia textualmente la línea que indica cómo se pagó, por ejemplo 'Pago: Credito' o "
    "'Débito'. Si ninguna línea lo indica, usa null ('Pago Cuentas' no cuenta como medio de pago).",
}

# nombre: (inicio, fin, campos, descripción). Inicio y fin son fracciones del alto;
# los recortes se superponen un poco para no cortar líneas en el borde.
REGIONES = {
    "encabezado": (0.0, 0.55, ["nombre_emisor", "rut_emisor", "fecha", "folio_boleta"], "parte superior"),
    "pie": (0.45, 1.0, ["monto_total", "iva", "linea_pago"], "parte inferior"),
}


def limpiar_monto(valor):
    """Convierte '$12.990', '12.990,50' o '12990' en float. Deja pasar números y None."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        # Los pesos no tienen decimales: 44.39 es '44.390' con el separador de miles mal leído.
        if isinstance(valor, float) and valor != int(valor):
            return float(round(valor * 1000))
        return valor
    texto = "".join(c for c in str(valor) if c.isdigit() or c in ".,")
    if not texto:
        return None
    if "," in texto:  # formato chileno con decimales: 12.990,50
        texto = texto.replace(".", "").replace(",", ".")
    elif texto.count(".") > 1 or (texto.count(".") == 1 and len(texto.split(".")[1]) == 3):
        texto = texto.replace(".", "")  # separador de miles: 12.990
    return float(texto)


def limpiar_fecha(valor):
    """Extrae la fecha de textos como '2026-09-28 13:56:15' o '28-09-2026 HORA 13:54'."""
    if not valor:
        return None
    texto = str(valor)
    if m := re.search(r"(\d{4})-(\d{2})-(\d{2})", texto):
        return m.group(0)
    if m := re.search(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4}|\d{2})\b", texto):
        dia, mes, anio = m.groups()
        anio = anio if len(anio) == 4 else f"20{anio}"
        return f"{anio}-{int(mes):02d}-{int(dia):02d}"
    return None


def limpiar_folio(valor):
    """Deja solo los dígitos: 'BOLETA ELECTRONICA 280186' -> '280186', '36.675.529' -> '36675529'."""
    if not valor:
        return None
    digitos = re.sub(r"\D", "", str(valor))
    return digitos or None


def _digito_verificador(cuerpo: str) -> str:
    suma = sum(int(d) * f for d, f in zip(reversed(cuerpo), [2, 3, 4, 5, 6, 7] * 4))
    resto = 11 - suma % 11
    return {11: "0", 10: "K"}.get(resto, str(resto))


def limpiar_rut(valor):
    """Normaliza a 76.178.360-2 y descarta RUTs con dígito verificador inválido
    (un RUT inventado por el modelo es peor que un campo vacío)."""
    if not valor:
        return None
    texto = str(valor)
    # Tolera artefactos de impresión como '77.648.040,00-1' (decimales antes del guion)
    if m := re.search(r"(\d{1,2}(?:\.?\d{3}){2})(?:[,.]\d+)?\s*-\s*([\dkK])", texto):
        cuerpo, dv = re.sub(r"\D", "", m.group(1)), m.group(2).upper()
    else:
        limpio = re.sub(r"[^0-9kK]", "", texto).upper()
        if len(limpio) < 7:
            return None
        cuerpo, dv = limpio[:-1], limpio[-1]
    if not cuerpo.isdigit() or _digito_verificador(cuerpo) != dv:
        return None
    return f"{int(cuerpo):,}".replace(",", ".") + f"-{dv}"


def limpiar_medio_pago(valor):
    if not valor:
        return None
    texto = str(valor).lower()
    for patron, nombre in [
        (r"cr[eé]di", "Crédito"),
        (r"d[eé]bi", "Débito"),
        (r"efectivo", "Efectivo"),
        (r"transfer", "Transferencia"),
    ]:
        if re.search(patron, texto):
            return nombre
    return None


class DatosBoleta(BaseModel):
    # Los modelos chicos a veces devuelven el folio o el RUT como número.
    model_config = ConfigDict(coerce_numbers_to_str=True)

    nombre_emisor: str | None = Field(default=None, description="Razón social o comercio")
    rut_emisor: str | None = Field(
        default=None, description="RUT del emisor con puntos y guion, si aparece"
    )
    fecha: str | None = Field(default=None, description="Fecha de emisión en formato YYYY-MM-DD")
    monto_total: float | None = Field(
        default=None, description="Monto total pagado en pesos, solo el número"
    )
    iva: float | None = Field(default=None, description="IVA de la boleta en pesos, solo el número")
    folio_boleta: str | None = Field(default=None, description="Número de boleta, solo dígitos")
    medio_pago: str | None = Field(default=None, description="Efectivo, Débito, Crédito, etc.")
    linea_pago: str | None = Field(
        default=None, description="Texto exacto de la línea que indica el medio de pago, si existe"
    )

    @model_validator(mode="before")
    @classmethod
    def _medio_pago_desde_linea(cls, datos):
        """El medio de pago se deduce del texto copiado de la boleta, no de lo que el modelo
        'cree': así 'Pago Cuentas' no se convierte en 'Efectivo'."""
        if isinstance(datos, dict) and "linea_pago" in datos:
            datos = {**datos, "medio_pago": datos["linea_pago"]}
        return datos

    _norm_monto = field_validator("monto_total", mode="before")(limpiar_monto)
    _norm_iva = field_validator("iva", mode="before")(limpiar_monto)
    _norm_fecha = field_validator("fecha", mode="before")(limpiar_fecha)
    _norm_folio = field_validator("folio_boleta", mode="before")(limpiar_folio)
    _norm_rut = field_validator("rut_emisor", mode="before")(limpiar_rut)
    _norm_pago = field_validator("medio_pago", mode="before")(limpiar_medio_pago)


COLUMNAS = ["nombre_archivo", *DatosBoleta.model_fields.keys(), "revisar", "error"]


def listar_imagenes(carpeta: Path) -> list[Path]:
    return sorted(
        f for f in carpeta.iterdir() if f.is_file() and f.suffix.lower() in EXTENSIONES_VALIDAS
    )


def abrir_imagen(ruta: Path) -> Image.Image:
    """Abre la imagen corrigiendo la orientación EXIF."""
    with Image.open(ruta) as img:
        return ImageOps.exif_transpose(img).convert("RGB")


def _umbral_otsu(gris: np.ndarray) -> int:
    hist = np.bincount(gris.ravel(), minlength=256).astype(float)
    total, suma = hist.sum(), float((np.arange(256) * hist).sum())
    peso_fondo = suma_fondo = mejor = 0.0
    umbral = 0
    for t in range(256):
        peso_fondo += hist[t]
        if peso_fondo == 0:
            continue
        peso_frente = total - peso_fondo
        if peso_frente == 0:
            break
        suma_fondo += t * hist[t]
        varianza = (
            peso_fondo
            * peso_frente
            * (suma_fondo / peso_fondo - (suma - suma_fondo) / peso_frente) ** 2
        )
        if varianza > mejor:
            mejor, umbral = varianza, t
    return umbral


def recortar_boleta(img: Image.Image, margen: float = 0.03) -> Image.Image:
    """Recorta horizontalmente la boleta (papel claro sobre fondo oscuro), para descartar
    boletas vecinas, dedos y fondo. Si no logra detectarla con confianza, devuelve la imagen entera."""
    ancho, alto = img.size
    ancho_chico = 300
    chica = img.convert("L").resize((ancho_chico, max(1, int(ancho_chico * alto / ancho))))
    gris = np.asarray(chica)
    claro = gris > _umbral_otsu(gris)
    perfil = np.convolve(claro.mean(axis=0), np.ones(5) / 5, mode="same")
    if perfil.max() == 0:
        return img
    activo = perfil > 0.5 * perfil.max()

    tramos, inicio = [], None
    for x, valor in enumerate(activo):
        if valor and inicio is None:
            inicio = x
        if inicio is not None and (not valor or x == len(activo) - 1):
            tramos.append((inicio, x if not valor else x + 1))
            inicio = None
    if not tramos:
        return img

    ini, fin = max(tramos, key=lambda t: t[1] - t[0])
    fraccion = (fin - ini) / ancho_chico
    if fraccion < 0.2 or fraccion > 0.9:  # detección poco creíble: mejor no recortar
        return img
    x0 = max(0, int((ini / ancho_chico - margen) * ancho))
    x1 = min(ancho, int((fin / ancho_chico + margen) * ancho))
    return img.crop((x0, 0, x1, alto))


def revisar(datos: dict, hoy: date | None = None) -> list[str]:
    """Validaciones cruzadas: devuelve avisos de lo que conviene revisar a mano."""
    hoy = hoy or date.today()
    avisos: list[str] = []

    monto, iva = datos.get("monto_total"), datos.get("iva")
    if monto is None:
        avisos.append("sin monto")
    elif iva is None:
        avisos.append("monto sin verificar (falta IVA)")
    else:
        esperado = round(monto * 19 / 119)  # en Chile IVA = total * 19/119
        if abs(esperado - iva) > 2:
            avisos.append(f"monto no cuadra con IVA (IVA esperado ~{esperado}, leído {iva:.0f})")

    fecha = datos.get("fecha")
    if not fecha:
        avisos.append("sin fecha")
    else:
        try:
            d = date.fromisoformat(fecha)
            if d > hoy or d < hoy - timedelta(days=400):
                avisos.append(f"fecha sospechosa ({fecha})")
        except ValueError:
            avisos.append(f"fecha inválida ({fecha})")

    if not datos.get("rut_emisor"):
        avisos.append("RUT vacío o inválido")
    if not datos.get("folio_boleta"):
        avisos.append("sin folio")
    return avisos


def recortar_a_jpeg(img: Image.Image, inicio: float, fin: float, max_lado: int = MAX_LADO_PX) -> bytes:
    """Recorta una franja vertical, mejora el contraste y devuelve JPEG en bytes."""
    ancho, alto = img.size
    franja = img.crop((0, int(alto * inicio), ancho, int(alto * fin)))
    franja = ImageOps.autocontrast(franja, cutoff=1)
    franja.thumbnail((max_lado, max_lado))
    buffer = io.BytesIO()
    franja.save(buffer, format="JPEG", quality=90)
    return buffer.getvalue()


def _llamar_modelo(modelo: str, imagen: bytes, formato, prompt: str) -> str:
    respuesta = ollama.chat(
        model=modelo,
        messages=[{"role": "user", "content": prompt, "images": [imagen]}],
        format=formato,
        options=OPCIONES_GENERACION,
        keep_alive="10m",
    )
    return respuesta["message"]["content"] or ""


def _extraer_json(texto: str) -> str:
    """Recorta desde la primera '{' hasta la última '}' (por si el modelo agrega texto)."""
    inicio, fin = texto.find("{"), texto.rfind("}")
    return texto[inicio : fin + 1] if inicio != -1 and fin > inicio else texto


def _esquema_para(campos: list[str]) -> dict:
    esquema = DatosBoleta.model_json_schema()
    esquema["properties"] = {k: v for k, v in esquema["properties"].items() if k in campos}
    esquema.pop("required", None)
    return esquema


def _construir_prompt(campos: list[str], descripcion: str) -> str:
    lineas = "\n".join(f"- {c}: {PISTAS[c]}" for c in campos)
    return (
        f"Esta imagen muestra la {descripcion} de una boleta electrónica chilena. "
        f"Extrae solo estos campos en JSON:\n{lineas}\n"
        "Ignora cualquier otra boleta que se vea parcialmente en los bordes. "
        "Copia los números exactamente como aparecen. "
        "Si un dato no es visible o no existe en esta imagen, usa null."
    )


def extraer_region(imagen: bytes, campos: list[str], descripcion: str, modelo: str) -> tuple[dict, str]:
    """Extrae los campos de un recorte. Si el modelo responde vacío o falla con el esquema,
    reintenta con modos más simples (algunos modelos chicos fallan con el esquema completo)."""
    prompt = _construir_prompt(campos, descripcion)
    prompt_json = f"{prompt}\nResponde solo con un objeto JSON con exactamente estas claves: {', '.join(campos)}."
    intentos = [
        ("esquema", _esquema_para(campos), prompt),
        ("json", "json", prompt_json),
        ("libre", None, prompt_json),
    ]

    ultimo_error = "sin respuesta"
    for nombre, formato, texto in intentos:
        try:
            contenido = _llamar_modelo(modelo, imagen, formato, texto)
        except ollama.ResponseError as e:
            ultimo_error = f"modo '{nombre}': {e.error}"
            continue
        if not contenido.strip():
            ultimo_error = f"respuesta vacía en modo '{nombre}'"
            continue
        try:
            datos = DatosBoleta.model_validate_json(_extraer_json(contenido))
            incluir = set(campos) | ({"medio_pago"} if "linea_pago" in campos else set())
            return datos.model_dump(include=incluir), contenido
        except ValidationError as e:
            detalle = "; ".join(
                f"{'.'.join(map(str, err['loc']))}: {err['msg']}" for err in e.errors()
            )
            ultimo_error = f"{detalle} | modo '{nombre}', respuesta: {contenido[:300]!r}"
    raise ValueError(ultimo_error)


def extraer_boleta(
    ruta: Path, modelo: str, autorecorte: bool = True
) -> tuple[dict, dict[str, str], list[str]]:
    """Procesa la boleta por regiones. Devuelve (datos, respuestas crudas, errores).
    Si una región falla, se conservan los datos de las demás."""
    img = abrir_imagen(ruta)
    if autorecorte:
        img = recortar_boleta(img)
    datos: dict = {}
    crudos: dict[str, str] = {}
    errores: list[str] = []
    for nombre, (inicio, fin, campos, descripcion) in REGIONES.items():
        try:
            resultado, crudo = extraer_region(recortar_a_jpeg(img, inicio, fin), campos, descripcion, modelo)
            datos.update(resultado)
            crudos[nombre] = crudo
        except ValueError as e:
            errores.append(f"{nombre}: {e}")
    return datos, crudos, errores


def procesar_carpeta(
    carpeta: Path, archivo_excel: Path, modelo: str, debug: bool = False, autorecorte: bool = True
) -> int:
    if not carpeta.is_dir():
        print(f"Error: la carpeta '{carpeta}' no existe.")
        return 1

    imagenes = listar_imagenes(carpeta)
    if not imagenes:
        print(f"No se encontraron imágenes ({', '.join(sorted(EXTENSIONES_VALIDAS))}) en '{carpeta}'.")
        return 1

    print(f"Procesando {len(imagenes)} boletas con '{modelo}' (2 recortes por boleta)...\n")

    registros: list[dict] = []
    for i, imagen in enumerate(imagenes, start=1):
        print(f"[{i}/{len(imagenes)}] {imagen.name} ... ", end="", flush=True)
        registro: dict = {"nombre_archivo": imagen.name, "error": None}
        try:
            datos, crudos, errores = extraer_boleta(imagen, modelo, autorecorte)
            registro.update(datos)
            registro["revisar"] = "; ".join(revisar(datos)) or None
            if errores:
                registro["error"] = " | ".join(errores)
            estado = "OK" if not errores else f"CON ERRORES: {registro['error']}"
            if registro["revisar"]:
                estado += f"  ⚠ revisar: {registro['revisar']}"
            print(estado)
            if debug:
                for nombre, crudo in crudos.items():
                    print(f"    [{nombre}] respuesta cruda: {crudo.strip()}")
        except (ollama.ResponseError, OSError) as e:
            registro["error"] = str(e)
            print(f"ERROR: {registro['error']}")
        registros.append(registro)

    df = pd.DataFrame(registros).reindex(columns=COLUMNAS)
    for columna in ("error", "revisar"):
        if df[columna].isna().all():
            df = df.drop(columns=columna)
    df.to_excel(archivo_excel, index=False)

    sin_errores = int(df["error"].isna().sum()) if "error" in df.columns else len(df)
    con_avisos = int(df["revisar"].notna().sum()) if "revisar" in df.columns else 0
    print(f"\n{sin_errores}/{len(imagenes)} boletas procesadas sin errores; {con_avisos} con avisos para revisar.")
    print(f"Archivo generado: {archivo_excel.resolve()}")
    return 0 if sin_errores > 0 else 1


def main() -> int:
    parser = argparse.ArgumentParser(description="Extrae datos de boletas con Ollama.")
    parser.add_argument("--carpeta", type=Path, default=Path("boletas"))
    parser.add_argument("--salida", type=Path, default=Path("boletas_extraidas.xlsx"))
    parser.add_argument("--modelo", default=MODELO_POR_DEFECTO)
    parser.add_argument(
        "--debug", action="store_true", help="muestra la respuesta cruda del modelo por recorte"
    )
    parser.add_argument(
        "--sin-autorecorte",
        action="store_true",
        help="no recortar la boleta dentro de la foto (útil si la detección falla)",
    )
    args = parser.parse_args()
    return procesar_carpeta(
        args.carpeta, args.salida, args.modelo, args.debug, autorecorte=not args.sin_autorecorte
    )


if __name__ == "__main__":
    sys.exit(main())