# OCR de boletas chilenas con un modelo de visión local

Extrae datos estructurados (emisor, RUT, fecha, folio, monto, IVA y medio de pago) desde **fotos de boletas electrónicas chilenas** y los guarda en un Excel. Todo corre en tu computador con [Ollama](https://ollama.com): las imágenes **no salen de tu máquina**, algo importante porque una boleta contiene RUT y datos de compra.

Está pensado para un notebook **sin GPU y con 8 GB de RAM**.

---

## Contenido

- [Qué hace](#qué-hace)
- [Requisitos](#requisitos)
- [Instalación](#instalación)
- [Uso](#uso)
- [Resultado](#resultado)
- [Cómo funciona](#cómo-funciona)
- [Validaciones y la columna `revisar`](#validaciones-y-la-columna-revisar)
- [Modelos probados](#modelos-probados)
- [Consejos para mejores fotos](#consejos-para-mejores-fotos)
- [Personalización](#personalización)
- [Solución de problemas](#solución-de-problemas)
- [Limitaciones](#limitaciones)
- [Estructura del proyecto](#estructura-del-proyecto)

---

## Qué hace

1. Lee todas las imágenes de una carpeta (`.jpg`, `.jpeg`, `.png`, `.webp`).
2. Detecta y recorta la boleta dentro de la foto (descarta dedos, fondo y boletas vecinas).
3. Divide la boleta en dos franjas (encabezado y pie) y le pide al modelo de visión que extraiga los campos de cada una.
4. Limpia y normaliza los valores (fechas, montos, RUT, folio, medio de pago).
5. Ejecuta validaciones cruzadas y marca lo que conviene revisar a mano.
6. Genera un archivo Excel con una fila por boleta.

Campos extraídos:

| Campo           | Descripción                                                                 |
| --------------- | --------------------------------------------------------------------------- |
| `nombre_emisor` | Razón social del comercio                                                   |
| `rut_emisor`    | RUT con formato `76.178.360-2` (se valida el dígito verificador)            |
| `fecha`         | Fecha de emisión en formato `YYYY-MM-DD`                                    |
| `monto_total`   | Total pagado en pesos                                                       |
| `iva`           | IVA de la boleta en pesos (se usa para verificar el monto)                  |
| `folio_boleta`  | Número de boleta, solo dígitos                                              |
| `medio_pago`    | Efectivo, Débito, Crédito o Transferencia (vacío si la boleta no lo indica) |
| `linea_pago`    | Línea de texto de la boleta de la que se dedujo el medio de pago            |

---

## Requisitos

- **Windows, macOS o Linux** (desarrollado y probado en Windows).
- **Python 3.10 o superior.**
- **[uv](https://docs.astral.sh/uv/)** como gestor de paquetes (o `pip`, ver más abajo).
- **[Ollama](https://ollama.com/download)** instalado y en ejecución.
- **RAM:** al menos 8 GB. El modelo por defecto (`gemma3:4b`) ocupa unos pocos GB; cierra el navegador y otros programas pesados mientras corre.
- **GPU:** no es necesaria. Sin GPU el procesamiento es más lento (ver [Limitaciones](#limitaciones)).

---

## Instalación

### 1. Instalar Ollama y descargar el modelo

Instala Ollama desde [ollama.com/download](https://ollama.com/download) y verifica que funciona:

```powershell
ollama --version
```

Descarga el modelo por defecto:

```powershell
ollama pull gemma3:4b
```

### 2. Instalar las dependencias de Python

Con `uv`, desde la carpeta del proyecto:

```powershell
uv add ollama pandas pydantic openpyxl pillow
```

O con `pip`:

```powershell
pip install ollama pandas pydantic openpyxl pillow
```

> `numpy` se instala solo como dependencia de `pandas`.

### 3. Preparar las imágenes

Crea una carpeta `boletas/` junto a `main.py` y copia ahí las fotos:

```
ocr-boletas/
├── main.py
└── boletas/
    ├── foto1.jpeg
    └── foto2.jpeg
```

---

## Uso

Asegúrate de que Ollama esté corriendo (normalmente se inicia solo; si no, abre la aplicación o ejecuta `ollama serve`). Luego:

```powershell
uv run main.py
```

Con `pip`, usa `python main.py`.

### Opciones

| Opción              | Valor por defecto        | Descripción                                           |
| ------------------- | ------------------------ | ----------------------------------------------------- |
| `--carpeta`         | `boletas`                | Carpeta con las imágenes                              |
| `--salida`          | `boletas_extraidas.xlsx` | Archivo Excel de salida                               |
| `--modelo`          | `gemma3:4b`              | Modelo de visión de Ollama                            |
| `--debug`           | desactivado              | Muestra la respuesta cruda del modelo por cada franja |
| `--sin-autorecorte` | desactivado              | No recorta la boleta dentro de la foto                |

### Ejemplos

```powershell
# Carpeta y archivo de salida personalizados
uv run main.py --carpeta C:\Documentos\boletas_septiembre --salida septiembre.xlsx

# Ver qué responde el modelo (útil para diagnosticar errores)
uv run main.py --debug

# Probar otro modelo
uv run main.py --modelo qwen2.5vl:7b

# Si el recorte automático falla con tus fotos
uv run main.py --sin-autorecorte
```

### Salida en consola

```
Procesando 2 boletas con 'gemma3:4b' (2 recortes por boleta)...

[1/2] foto1.jpeg ... OK
[2/2] foto2.jpeg ... OK  ⚠ revisar: monto no cuadra con IVA (IVA esperado ~7087, leído 701)

2/2 boletas procesadas sin errores; 1 con avisos para revisar.
Archivo generado: C:\...\boletas_extraidas.xlsx
```

El programa termina con código de error si **ninguna** boleta se procesó sin errores.

---

## Resultado

El Excel tiene una fila por boleta, con estas columnas:

`nombre_archivo`, `nombre_emisor`, `rut_emisor`, `fecha`, `monto_total`, `iva`, `folio_boleta`, `medio_pago`, `linea_pago`, `revisar`, `error`

- **`revisar`**: avisos de las validaciones cruzadas (vacía si no hay nada que revisar). Si ninguna boleta tiene avisos, la columna no aparece.
- **`error`**: errores técnicos al procesar (por ejemplo, el modelo no respondió). Si no hubo errores, la columna no aparece.

> **Atención:** el programa **sobrescribe** el archivo de salida en cada ejecución. Si quieres conservar resultados anteriores, cambia el nombre con `--salida`.

---

## Cómo funciona

```
foto ─► autorecorte ─► franja superior (0–55%) ─► modelo ─► emisor, RUT, fecha, folio ─┐
         de la boleta └► franja inferior (45–100%) ─► modelo ─► monto, IVA, línea de pago ─┤
                                                                                        ▼
                                              normalización ─► validaciones ─► fila del Excel
```

### Por qué dos franjas

Los modelos de visión pequeños reducen la imagen a una resolución fija. Con una foto vertical completa, el texto chico del encabezado (RUT, fecha, folio) se vuelve ilegible y el modelo **inventa** valores plausibles. Dividir la boleta en dos franjas, con una pequeña superposición, hace que el texto ocupe más espacio en lo que ve el modelo. Cada franja usa además un esquema JSON con solo sus campos, lo que reduce las alucinaciones.

### Autorecorte horizontal

Las boletas son papel claro sobre un fondo oscuro. El programa calcula un umbral (método de Otsu) sobre una versión reducida de la foto, busca el tramo de columnas más ancho que sea claro y recorta ahí. Así se descartan dedos, la mesa y boletas vecinas, que de otro modo contaminan la lectura (por ejemplo, tomar el `POS: 134` de la boleta de al lado como folio).

Si la detección no es creíble (el tramo ocupa menos del 20% o más del 90% del ancho), no recorta y usa la imagen completa.

### Preparación de la imagen

Cada franja se corrige en orientación (EXIF), se le aumenta el contraste y se reduce a un máximo de 1280 px por lado, lo que acelera el procesamiento en CPU.

### Reintentos por modo de salida

Algunos modelos pequeños fallan con la salida estructurada completa (responden vacío o con errores). Por eso cada franja se intenta en tres modos, en orden, hasta que uno funcione:

1. **`esquema`**: salida estructurada con el esquema JSON de los campos de esa franja.
2. **`json`**: `format="json"` con las claves listadas en el prompt.
3. **`libre`**: sin formato, recortando el JSON entre la primera `{` y la última `}`.

Si una franja falla, se conservan los datos de la otra y el error queda en la columna `error`.

### Opciones de generación

Parámetros enviados a Ollama (definidos en `OPCIONES_GENERACION`):

| Parámetro        | Valor  | Motivo                                                               |
| ---------------- | ------ | -------------------------------------------------------------------- |
| `temperature`    | `0.1`  | Con `0` (greedy) los modelos pequeños caen en bucles de repetición   |
| `repeat_penalty` | `1.1`  | Reduce esos bucles                                                   |
| `num_predict`    | `512`  | El JSON de una boleta cabe de sobra; corta respuestas descontroladas |
| `num_ctx`        | `4096` | Contexto acotado: menos RAM y más velocidad en CPU                   |

El modelo se mantiene cargado 10 minutos (`keep_alive`) para no recargarlo entre boletas.

### El medio de pago se deduce de evidencia

Pedirle directamente el medio de pago a un modelo pequeño lo hacía inventar valores (por ejemplo, "Efectivo" cuando la boleta solo decía "Pago Cuentas"). Ahora se le pide que **copie textualmente la línea** donde se indica cómo se pagó (`linea_pago`), y el medio se deduce de esa línea buscando palabras clave (crédito, débito, efectivo, transferencia). Si la línea no contiene ninguna, el campo queda vacío.

---

## Validaciones y la columna `revisar`

Tras extraer los datos se normalizan y se verifican. Los valores que no se pueden normalizar quedan vacíos, y los que son sospechosos generan un aviso en `revisar`.

### Normalización

| Campo         | Qué hace                                                                                                                                                                                                 |
| ------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Monto / IVA   | Acepta `$12.990`, `12.990,50`, `12990`. Los pesos no tienen decimales: un valor como `44.39` se interpreta como `44.390` (separador de miles) y se corrige a `44390`                                     |
| Fecha         | Extrae la fecha de textos como `26-09-2026 HORA : 13:54:00` y la convierte a `YYYY-MM-DD`. Acepta separadores `-`, `/`, `.` y años de dos dígitos                                                        |
| Folio         | Deja solo los dígitos (`36.675.529` → `36675529`)                                                                                                                                                        |
| RUT           | Normaliza a `76.178.360-2`, tolera artefactos de impresión como `77.648.040,00-1` y **descarta** los RUT con dígito verificador inválido (módulo 11), porque un RUT inventado es peor que un campo vacío |
| Medio de pago | Se reduce a Crédito, Débito, Efectivo o Transferencia                                                                                                                                                    |
| Emisor        | Colapsa saltos de línea y espacios repetidos                                                                                                                                                             |

### Avisos en `revisar`

| Aviso                             | Cuándo aparece                                                                                        |
| --------------------------------- | ----------------------------------------------------------------------------------------------------- |
| `monto no cuadra con IVA`         | El IVA esperado (`total × 19/119`) difiere en más de 2 pesos del IVA leído. Detecta montos mal leídos |
| `monto sin verificar (falta IVA)` | Hay monto pero no IVA para contrastarlo                                                               |
| `sin monto`                       | No se extrajo el monto                                                                                |
| `fecha sospechosa`                | La fecha es futura o tiene más de 400 días de antigüedad                                              |
| `sin fecha` / `fecha inválida`    | No hay fecha o no es una fecha real                                                                   |
| `RUT vacío o inválido`            | No hay RUT o su dígito verificador no cuadra                                                          |
| `sin folio`                       | No se extrajo el folio                                                                                |

**Flujo recomendado:** abre el Excel y filtra primero las filas con `revisar` no vacío. Verifica contra la boleta original una muestra de las demás, sobre todo el **folio**, que ninguna validación puede comprobar.

---

## Modelos probados

Probado con **Ollama 0.35.0 en Windows**, sin GPU:

| Modelo            | Resultado                                                                                                                |
| ----------------- | ------------------------------------------------------------------------------------------------------------------------ |
| `gemma3:4b`       | ✅ Funciona. Es el modelo por defecto                                                                                    |
| `qwen2.5vl:3b`    | ❌ Falló en las pruebas: respuestas vacías con salida estructurada y bucles de repetición (`token repeat limit reached`) |
| `llama3.2-vision` | ❌ No carga: `unknown model architecture: 'mllama'`                                                                      |

### Sobre `llama3.2-vision`

Ese error no depende de tu instalación ni de este código. Es un problema reportado en el repositorio de Ollama ([issue #16547](https://github.com/ollama/ollama/issues/16547)) que afecta a versiones recientes, y reinstalar o volver a descargar el modelo no lo soluciona. Mientras no se corrija, usa otro modelo.

### Otros modelos

Puedes probar cualquier modelo de visión de Ollama con `--modelo`. Ten en cuenta:

- Con 8 GB de RAM, evita modelos de 7B o más: pueden quedarse sin memoria o volverse muy lentos.
- Los modelos más grandes suelen leer mejor los números pequeños, pero necesitan más RAM.
- Si un modelo responde vacío o entra en bucles, usa `--debug` para ver qué devuelve.

---

## Consejos para mejores fotos

La calidad de la foto pesa más que cualquier ajuste del programa:

- **Una sola boleta por foto.** Si hay otra a un costado, puede contaminar la lectura aunque el autorecorte intente descartarla.
- **Fondo oscuro y papel claro.** El autorecorte funciona por contraste.
- **Boleta lo más recta y plana posible.** Alisa los pliegues, sobre todo en las líneas de encabezado y total.
- **Buena luz y sin sombras** sobre el texto. Evita el flash directo, que genera reflejos en el papel térmico.
- **Que la boleta ocupe buena parte del cuadro** y salga nítida.
- **Sin dedos tapando texto.**

Las fotos enviadas por WhatsApp llegan comprimidas; si es posible, usa la foto original de la cámara.

---

## Personalización

Todo está en `main.py`.

### Cambiar o agregar campos

1. Agrega el campo a la clase `DatosBoleta` (con su descripción).
2. Agrega una pista en el diccionario `PISTAS` explicando dónde buscarlo en la boleta.
3. Asígnalo a una franja en `REGIONES`.
4. Si necesita limpieza, escribe una función `limpiar_*` y regístrala con `field_validator`.

La columna aparece sola en el Excel, porque `COLUMNAS` se construye a partir del modelo.

### Ajustar las franjas

`REGIONES` define, para cada franja, el inicio y fin como fracción del alto de la boleta, los campos que se le piden y una descripción. Si tus boletas son muy largas (por ejemplo, de supermercado con muchos ítems), puede convenir reducir la franja superior o agregar una tercera.

### Ajustar las validaciones

En la función `revisar`:

- La tolerancia del IVA es de 2 pesos (`abs(esperado - iva) > 2`).
- La ventana de fechas es de 400 días (`timedelta(days=400)`).

### Usar en otro país

Las validaciones son específicas de Chile: RUT con dígito verificador módulo 11, IVA del 19% y formato de montos en pesos sin decimales. Para otro país hay que adaptar `limpiar_rut`, `limpiar_monto` y la verificación del IVA en `revisar`.

---

## Solución de problemas

| Síntoma                                             | Causa probable                                      | Qué hacer                                                                          |
| --------------------------------------------------- | --------------------------------------------------- | ---------------------------------------------------------------------------------- |
| `ConnectionError` o no conecta con Ollama           | Ollama no está corriendo                            | Abre la aplicación de Ollama o ejecuta `ollama serve`                              |
| `model '...' not found`                             | El modelo no está descargado                        | `ollama pull gemma3:4b`                                                            |
| `unknown model architecture: 'mllama'`              | Bug de Ollama con `llama3.2-vision`                 | Usa otro modelo (ver [Modelos probados](#modelos-probados))                        |
| `Invalid JSON: EOF while parsing` / respuesta vacía | El modelo no generó nada con la salida estructurada | El programa reintenta con otros modos. Si persiste, prueba otro modelo             |
| `prediction aborted, token repeat limit reached`    | El modelo cayó en un bucle de repetición            | Ya se mitiga con `repeat_penalty` y `temperature`. Si persiste, prueba otro modelo |
| Muy lento o el equipo se congela                    | Poca RAM disponible                                 | Cierra otros programas, usa un modelo más pequeño                                  |
| Campos vacíos en muchas boletas                     | Foto borrosa, torcida o con poca luz                | Revisa [Consejos para mejores fotos](#consejos-para-mejores-fotos) y usa `--debug` |
| Folio o fecha tomados de otra boleta                | Había una boleta vecina en la foto                  | Una sola boleta por foto, o recórtala antes                                        |
| El recorte automático corta mal                     | Contraste insuficiente entre papel y fondo          | Usa `--sin-autorecorte`                                                            |
| `PermissionError` al guardar el Excel               | El archivo está abierto en Excel                    | Ciérralo y vuelve a ejecutar                                                       |
| `No se encontraron imágenes`                        | La carpeta está vacía o con otras extensiones       | Revisa `--carpeta` y que las fotos sean `.jpg`, `.jpeg`, `.png` o `.webp`          |

### Diagnosticar con `--debug`

```powershell
uv run main.py --debug
```

Muestra, por cada franja, la respuesta cruda del modelo **antes** de la normalización. Con eso se distingue si el error es del modelo (leyó mal) o de la limpieza posterior.

---

## Limitaciones

- **Probado con pocas boletas.** El desarrollo se validó con dos boletas reales. Antes de procesar un lote grande, prueba con 10 o 15 variadas (distintos comercios, fotos torcidas, poca luz, boletas largas) para conocer la precisión real.
- **Un modelo de 4B se equivoca.** Puede leer mal dígitos pequeños o inventar datos que no están.
- **El folio no se puede verificar.** Ninguna validación detecta un folio equivocado. Revisa una muestra contra el original, sobre todo si lo usarás para contabilidad.
- **El medio de pago puede seguir siendo erróneo.** Se deduce de la línea que el modelo dice haber copiado; si inventa la línea completa, el resultado será incorrecto.
- **Boletas largas.** Las franjas se cortan por porcentaje fijo. En una boleta con muchos ítems, el monto puede quedar fuera de la franja inferior o el encabezado puede quedar partido.
- **Velocidad en CPU.** Son dos llamadas al modelo por boleta; para lotes grandes conviene dejar el proceso corriendo.
- **Pensado para boletas electrónicas chilenas.** Facturas, boletas manuscritas u otros formatos pueden requerir ajustar prompts y validaciones.
- **Si necesitas más confiabilidad:** una alternativa es combinar un OCR clásico (Tesseract, PaddleOCR) con expresiones regulares, ya que el formato de estas boletas es muy regular. Suele leer mejor los números pequeños que un modelo de visión pequeño.

---

## Estructura del proyecto

```
ocr-boletas/
├── main.py                  # Todo el programa
├── README.md                # Este archivo
├── boletas/                 # Imágenes de entrada
└── boletas_extraidas.xlsx   # Resultado (se genera al ejecutar)
```

Organización de `main.py`:

| Elemento                   | Función                                                          |
| -------------------------- | ---------------------------------------------------------------- |
| `PISTAS`, `REGIONES`       | Instrucciones por campo y definición de las franjas              |
| `limpiar_*`                | Normalización de monto, fecha, folio, RUT, medio de pago y texto |
| `DatosBoleta`              | Esquema Pydantic con los campos y sus validadores                |
| `recortar_boleta`          | Detección y recorte horizontal de la boleta                      |
| `recortar_a_jpeg`          | Recorte vertical, contraste y reducción de la franja             |
| `extraer_region`           | Llamada al modelo con reintentos por modo de salida              |
| `extraer_boleta`           | Procesa una boleta completa por franjas                          |
| `revisar`                  | Validaciones cruzadas y avisos                                   |
| `procesar_carpeta`, `main` | Recorrido de la carpeta, Excel de salida y línea de comandos     |
