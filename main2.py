"""Extracción de datos de boletas chilenas con la API de NVIDIA NIM (modelos de visión alojados).

Versión solo-NVIDIA de main.py. La clave se lee desde un archivo .env:

    # .env  (junto a este archivo; NO lo subas a git)
    NVIDIA_API_KEY=nvapi-...
    NVIDIA_MODEL=meta/llama-3.2-90b-vision-instruct   # opcional

Instalación:
    uv add openai python-dotenv pandas pydantic openpyxl pillow

Uso:
    uv run main2.py --listar-modelos          # modelos de visión disponibles con tu clave
    uv run main2.py                           # elige solo un modelo disponible
    uv run main2.py --modelo meta/llama-3.2-11b-vision-instruct --debug

AVISO DE PRIVACIDAD: cada imagen (con RUT, comercio y montos) se envía a los servidores de NVIDIA.
"""

from __future__ import annotations

import argparse
import base64
import io
import os
import re
import sys
import time
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
from PIL import Image, ImageOps
from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    ValidationError,
    field_validator,
    model_validator,
)

try:
    import openai
    from dotenv import load_dotenv
except ImportError as e:  # mensaje claro en vez de un traceback
    sys.exit(
        f"Falta el paquete '{e.name}'. Instala las dependencias con:\n"
        "  uv add openai python-dotenv pandas pydantic openpyxl pillow"
    )

# Modelos de visión en orden de preferencia: se usa el primero que aparezca disponible con tu clave.
# Llama 3.2 Vision tiene salida estructurada documentada por NVIDIA y está en la capa gratuita
# (según listados de terceros de octubre de 2026). Los modelos "reasoning" se evitan a propósito:
# gastan tokens pensando y pueden romper el JSON.
MODELOS_PREFERIDOS = [
    "meta/llama-3.2-90b-vision-instruct",
    "meta/llama-3.2-11b-vision-instruct",
]
NIM_BASE_URL = "https://integrate.api.nvidia.com/v1"
# Precaución: algunos endpoints alojados de NVIDIA limitan la imagen en base64 a ~180.000
# caracteres (~135 KB). No lo confirmé para los modelos de visión; --limite-imagen-kb 0 lo desactiva.
LIMITE_IMAGEN_KB = 135
# La capa gratuita se reporta en ~40 solicitudes/minuto (fuentes de terceros); por defecto se
# usan 30 para dejar margen. Cada boleta consume 2 solicitudes.
RPM_POR_DEFECTO = 30
PALABRAS_VISION = (
    "vl",
    "vision",
    "vila",
    "pixtral",
    "paligemma",
    "phi-3.5",
    "phi-4-multimodal",
    "gemma",
    "ministral",
    "llava",
    "qwen",
    "omni",
    "glimmer",
)
EXTENSIONES_VALIDAS = {".jpg", ".jpeg", ".png", ".webp"}
MAX_LADO_PX = 1280

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
    "encabezado": (
        0.0,
        0.55,
        ["nombre_emisor", "rut_emisor", "fecha", "folio_boleta"],
        "parte superior",
    ),
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
    elif texto.count(".") > 1 or (
        texto.count(".") == 1 and len(texto.split(".")[1]) == 3
    ):
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


def limpiar_texto(valor):
    """Colapsa saltos de línea y espacios repetidos ('A\nB' -> 'A B')."""
    if not valor:
        return None
    return " ".join(str(valor).split()) or None


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

    nombre_emisor: str | None = Field(
        default=None, description="Razón social o comercio"
    )
    rut_emisor: str | None = Field(
        default=None, description="RUT del emisor con puntos y guion, si aparece"
    )
    fecha: str | None = Field(
        default=None, description="Fecha de emisión en formato YYYY-MM-DD"
    )
    monto_total: float | None = Field(
        default=None, description="Monto total pagado en pesos, solo el número"
    )
    iva: float | None = Field(
        default=None, description="IVA de la boleta en pesos, solo el número"
    )
    folio_boleta: str | None = Field(
        default=None, description="Número de boleta, solo dígitos"
    )
    medio_pago: str | None = Field(
        default=None, description="Efectivo, Débito, Crédito, etc."
    )
    linea_pago: str | None = Field(
        default=None,
        description="Texto exacto de la línea que indica el medio de pago, si existe",
    )

    @model_validator(mode="before")
    @classmethod
    def _medio_pago_desde_linea(cls, datos):
        """El medio de pago se deduce del texto copiado de la boleta, no de lo que el modelo
        'cree': así 'Pago Cuentas' no se convierte en 'Efectivo'."""
        if isinstance(datos, dict) and "linea_pago" in datos:
            datos = {**datos, "medio_pago": datos["linea_pago"]}
        return datos

    _norm_emisor = field_validator("nombre_emisor", "linea_pago", mode="before")(
        limpiar_texto
    )
    _norm_monto = field_validator("monto_total", mode="before")(limpiar_monto)
    _norm_iva = field_validator("iva", mode="before")(limpiar_monto)
    _norm_fecha = field_validator("fecha", mode="before")(limpiar_fecha)
    _norm_folio = field_validator("folio_boleta", mode="before")(limpiar_folio)
    _norm_rut = field_validator("rut_emisor", mode="before")(limpiar_rut)
    _norm_pago = field_validator("medio_pago", mode="before")(limpiar_medio_pago)


COLUMNAS = ["nombre_archivo", *DatosBoleta.model_fields.keys(), "revisar", "error"]


def listar_imagenes(carpeta: Path) -> list[Path]:
    return sorted(
        f
        for f in carpeta.iterdir()
        if f.is_file() and f.suffix.lower() in EXTENSIONES_VALIDAS
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
    chica = img.convert("L").resize(
        (ancho_chico, max(1, int(ancho_chico * alto / ancho)))
    )
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
            avisos.append(
                f"monto no cuadra con IVA (IVA esperado ~{esperado}, leído {iva:.0f})"
            )

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


def _jpeg(img: Image.Image, calidad: int) -> bytes:
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=calidad)
    return buffer.getvalue()


def recortar_a_jpeg(
    img: Image.Image,
    inicio: float,
    fin: float,
    max_lado: int = MAX_LADO_PX,
    max_bytes: int | None = None,
) -> bytes:
    """Recorta una franja vertical, mejora el contraste y devuelve JPEG en bytes.
    Si se indica max_bytes, baja la calidad (y si hace falta el tamaño) hasta que quepa."""
    ancho, alto = img.size
    franja = img.crop((0, int(alto * inicio), ancho, int(alto * fin)))
    franja = ImageOps.autocontrast(franja, cutoff=1)
    franja.thumbnail((max_lado, max_lado))
    datos = _jpeg(franja, 90)
    if not max_bytes:
        return datos
    while len(datos) > max_bytes:
        for calidad in (80, 70, 60):
            datos = _jpeg(franja, calidad)
            if len(datos) <= max_bytes:
                return datos
        nuevo = (int(franja.width * 0.85), int(franja.height * 0.85))
        if min(nuevo) < 300:  # no seguir achicando: ya sería ilegible
            break
        franja = franja.resize(nuevo, Image.LANCZOS)
        datos = _jpeg(franja, 70)
    return datos


class ErrorModelo(Exception):
    """Error al llamar al modelo; se puede reintentar con otro modo de salida."""


class ErrorFatal(Exception):
    """Error que no tiene sentido reintentar (clave inválida, modelo retirado, etc.)."""


def cargar_clave() -> str:
    """Lee NVIDIA_API_KEY desde .env (junto al script y en el directorio actual) o del entorno.
    Las variables ya definidas en el sistema tienen prioridad sobre las del .env."""
    load_dotenv(Path(__file__).resolve().with_name(".env"))
    load_dotenv()
    clave = os.environ.get("NVIDIA_API_KEY", "").strip().strip("\"'")
    if not clave:
        raise ErrorFatal(
            "No se encontró NVIDIA_API_KEY. Crea un archivo .env junto a main2.py con la línea:\n"
            "  NVIDIA_API_KEY=nvapi-...\n"
            "(la clave se genera en build.nvidia.com)"
        )
    if not clave.startswith("nvapi-"):
        print(
            "Aviso: la clave no empieza con 'nvapi-'; revisa que sea una clave de build.nvidia.com."
        )
    return clave


class ClienteNVIDIA:
    """Cliente de la API alojada de NVIDIA NIM (compatible con OpenAI)."""

    nombre = "nvidia"
    # esquema: response_format json_schema; guided: extensión nvext.guided_json de NVIDIA
    MODOS = ("esquema", "guided", "json", "libre")

    def __init__(self, rpm: int = RPM_POR_DEFECTO, base_url: str = NIM_BASE_URL):
        # max_retries: el SDK reintenta solo los 429 (límite de solicitudes) con espera creciente.
        self.cliente = openai.OpenAI(
            base_url=base_url, api_key=cargar_clave(), max_retries=5, timeout=120
        )
        self.modelo = ""
        self._intervalo = 60.0 / rpm if rpm > 0 else 0.0
        self._ultima = 0.0

    def listar_modelos(self) -> list[str]:
        try:
            return sorted(m.id for m in self.cliente.models.list().data)
        except (openai.AuthenticationError, openai.PermissionDeniedError) as e:
            raise ErrorFatal(
                f"NVIDIA rechazó la clave (¿expiró o está mal copiada?): {e}"
            ) from e
        except openai.APIError as e:
            raise ErrorFatal(
                f"No se pudo consultar la lista de modelos de NVIDIA: {e}"
            ) from e

    def seleccionar_modelo(self, pedido: str | None) -> None:
        """Usa el modelo pedido; si no hay, el primero de MODELOS_PREFERIDOS que esté disponible."""
        disponibles = self.listar_modelos()
        if pedido:
            if pedido not in disponibles:
                print(
                    f"Aviso: '{pedido}' no aparece en la lista de modelos de NVIDIA; se intenta igual."
                )
            self.modelo = pedido
            return
        for candidato in MODELOS_PREFERIDOS:
            if candidato in disponibles:
                self.modelo = candidato
                print(f"Modelo elegido automáticamente: {candidato}")
                return
        parecidos = [
            m for m in disponibles if any(p in m.lower() for p in PALABRAS_VISION)
        ]
        raise ErrorFatal(
            "Ninguno de los modelos preferidos está disponible con tu clave.\n"
            "Elige uno con --modelo o NVIDIA_MODEL. Modelos que parecen de visión:\n"
            + "\n".join(f"  {m}" for m in parecidos)
        )

    def _esperar_turno(self) -> None:
        falta = self._intervalo - (time.monotonic() - self._ultima)
        if falta > 0:
            time.sleep(falta)
        self._ultima = time.monotonic()

    def llamar(self, imagen: bytes, prompt: str, modo: str, esquema: dict) -> str:
        b64 = base64.b64encode(imagen).decode()
        contenido = [
            {
                "type": "image_url",
                "image_url": {"url": f"data:image/jpeg;base64,{b64}"},
            },
            {"type": "text", "text": prompt},
        ]
        extra: dict = {}
        if modo == "esquema":
            extra["response_format"] = {
                "type": "json_schema",
                "json_schema": {"name": "boleta", "schema": esquema},
            }
        elif modo == "guided":
            extra["extra_body"] = {"nvext": {"guided_json": esquema}}
        elif modo == "json":
            extra["response_format"] = {"type": "json_object"}

        self._esperar_turno()
        try:
            respuesta = self.cliente.chat.completions.create(
                model=self.modelo,
                messages=[{"role": "user", "content": contenido}],
                temperature=0.1,
                max_tokens=512,
                **extra,
            )
        except (
            openai.AuthenticationError,
            openai.PermissionDeniedError,
            openai.NotFoundError,
        ) as e:
            raise ErrorFatal(
                f"NVIDIA rechazó la solicitud (¿clave inválida o modelo '{self.modelo}' no disponible "
                f"para tu cuenta? Prueba --listar-modelos): {e}"
            ) from e
        except openai.APIStatusError as e:
            if e.status_code == 410:  # NVIDIA responde 410 Gone a los modelos retirados
                raise ErrorFatal(
                    f"El modelo '{self.modelo}' fue retirado por NVIDIA (HTTP 410). "
                    "Elige otro con --listar-modelos y --modelo."
                ) from e
            raise ErrorModelo(
                str(e)
            ) from e  # 400 por parámetro no soportado, 429 agotado, 5xx...
        except openai.APIError as e:  # conexión, timeout...
            raise ErrorModelo(str(e)) from e
        return respuesta.choices[0].message.content or ""


def _extraer_json(texto: str) -> str:
    """Recorta desde la primera '{' hasta la última '}' (por si el modelo agrega texto)."""
    inicio, fin = texto.find("{"), texto.rfind("}")
    return texto[inicio : fin + 1] if inicio != -1 and fin > inicio else texto


def _esquema_para(campos: list[str]) -> dict:
    esquema = DatosBoleta.model_json_schema()
    esquema["properties"] = {
        k: v for k, v in esquema["properties"].items() if k in campos
    }
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


def extraer_region(
    imagen: bytes, campos: list[str], descripcion: str, backend
) -> tuple[dict, str]:
    """Extrae los campos de un recorte. Si el modelo responde vacío o falla con el esquema,
    reintenta con modos más simples (algunos modelos chicos fallan con el esquema completo)."""
    prompt = _construir_prompt(campos, descripcion)
    prompt_json = f"{prompt}\nResponde solo con un objeto JSON con exactamente estas claves: {', '.join(campos)}."
    esquema = _esquema_para(campos)

    ultimo_error = "sin respuesta"
    for nombre in backend.MODOS:
        texto = prompt if nombre in ("esquema", "guided") else prompt_json
        try:
            contenido = backend.llamar(imagen, texto, nombre, esquema)
        except ErrorModelo as e:
            ultimo_error = f"modo '{nombre}': {e}"
            continue
        if not contenido.strip():
            ultimo_error = f"respuesta vacía en modo '{nombre}'"
            continue
        try:
            datos = DatosBoleta.model_validate_json(_extraer_json(contenido))
            incluir = set(campos) | (
                {"medio_pago"} if "linea_pago" in campos else set()
            )
            return datos.model_dump(include=incluir), contenido
        except ValidationError as e:
            detalle = "; ".join(
                f"{'.'.join(map(str, err['loc']))}: {err['msg']}" for err in e.errors()
            )
            ultimo_error = (
                f"{detalle} | modo '{nombre}', respuesta: {contenido[:300]!r}"
            )
    raise ValueError(ultimo_error)


def extraer_boleta(
    ruta: Path, backend, autorecorte: bool = True, max_bytes: int | None = None
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
            franja = recortar_a_jpeg(img, inicio, fin, max_bytes=max_bytes)
            resultado, crudo = extraer_region(franja, campos, descripcion, backend)
            datos.update(resultado)
            crudos[nombre] = crudo
        except ValueError as e:
            errores.append(f"{nombre}: {e}")
    return datos, crudos, errores


def procesar_carpeta(
    carpeta: Path,
    archivo_excel: Path,
    backend,
    debug: bool = False,
    autorecorte: bool = True,
    max_bytes: int | None = None,
) -> int:
    if not carpeta.is_dir():
        print(f"Error: la carpeta '{carpeta}' no existe.")
        return 1

    imagenes = listar_imagenes(carpeta)
    if not imagenes:
        print(
            f"No se encontraron imágenes ({', '.join(sorted(EXTENSIONES_VALIDAS))}) en '{carpeta}'."
        )
        return 1

    print(
        f"Procesando {len(imagenes)} boletas con '{backend.modelo}' "
        f"[{backend.nombre}] (2 recortes por boleta)...\n"
    )
    print("Aviso: las imágenes se envían a los servidores de NVIDIA.\n")

    registros: list[dict] = []
    for i, imagen in enumerate(imagenes, start=1):
        print(f"[{i}/{len(imagenes)}] {imagen.name} ... ", end="", flush=True)
        registro: dict = {"nombre_archivo": imagen.name, "error": None}
        try:
            datos, crudos, errores = extraer_boleta(
                imagen, backend, autorecorte, max_bytes
            )
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
        except ErrorFatal as e:
            print(f"ERROR FATAL: {e}\nSe detiene el proceso; no se genera el Excel.")
            return 2
        except OSError as e:
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
    print(
        f"\n{sin_errores}/{len(imagenes)} boletas procesadas sin errores; {con_avisos} con avisos para revisar."
    )
    print(f"Archivo generado: {archivo_excel.resolve()}")
    return 0 if sin_errores > 0 else 1


def listar_modelos(cliente: ClienteNVIDIA, todos: bool) -> int:
    ids = cliente.listar_modelos()
    if not todos:
        ids = [i for i in ids if any(p in i.lower() for p in PALABRAS_VISION)]
    print(
        f"{len(ids)} modelos{'' if todos else ' (filtrados por nombre; puede faltar alguno de visión)'}:"
    )
    for i in ids:
        print(f"  {i}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extrae datos de boletas con la API de NVIDIA NIM."
    )
    parser.add_argument("--carpeta", type=Path, default=Path("boletas"))
    parser.add_argument("--salida", type=Path, default=Path("boletas_extraidas.xlsx"))
    parser.add_argument(
        "--modelo",
        default=None,
        help="modelo de visión; por defecto NVIDIA_MODEL del .env o el primero disponible de la lista preferida",
    )
    parser.add_argument(
        "--listar-modelos",
        nargs="?",
        const="vision",
        choices=["vision", "todos"],
        help="lista los modelos disponibles para tu clave y termina",
    )
    parser.add_argument(
        "--limite-imagen-kb",
        type=int,
        default=LIMITE_IMAGEN_KB,
        help=f"tamaño máximo de cada imagen enviada (por defecto {LIMITE_IMAGEN_KB}); 0 = sin límite",
    )
    parser.add_argument(
        "--rpm",
        type=int,
        default=RPM_POR_DEFECTO,
        help=f"máximo de solicitudes por minuto (por defecto {RPM_POR_DEFECTO}); 0 = sin pausa",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="muestra la respuesta cruda del modelo por recorte",
    )
    parser.add_argument(
        "--sin-autorecorte",
        action="store_true",
        help="no recortar la boleta dentro de la foto (útil si la detección falla)",
    )
    args = parser.parse_args()

    try:
        cliente = ClienteNVIDIA(rpm=args.rpm)
        if args.listar_modelos:
            return listar_modelos(cliente, todos=args.listar_modelos == "todos")
        cliente.seleccionar_modelo(
            args.modelo or os.environ.get("NVIDIA_MODEL") or None
        )
    except ErrorFatal as e:
        print(f"Error: {e}")
        return 2

    return procesar_carpeta(
        args.carpeta,
        args.salida,
        cliente,
        args.debug,
        autorecorte=not args.sin_autorecorte,
        max_bytes=args.limite_imagen_kb * 1024 or None,
    )


if __name__ == "__main__":
    sys.exit(main())
