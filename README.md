# md_converter.py

Convierte documentos (PDF, Word, Excel, PowerPoint, HTML, CSV e imágenes) a Markdown para pasárselos a una IA gastando menos tokens y obteniendo mejores resultados.

Está basado en [MarkItDown](https://github.com/microsoft/markitdown) de Microsoft y le agrega:

- Tablas de PDF con las celdas bien separadas (sin números pegados).
- Marcadores de página en los PDF y sin encabezados ni pies de página repetidos.
- Excel dividido en tablas (bloques), sin celdas "NaN" ni redondeos de valores.
- Descripción breve de las imágenes con IA (opcional), pensada para que la IA que lea el `.md` entienda el contexto y te pida la imagen original cuando necesite una cifra.

## Cuántos tokens ahorra

Comparado con la conversión estándar de MarkItDown (sin descripción de imágenes salvo donde se indica):

| Archivo | MarkItDown puro | md_converter | Cambio |
|---|---:|---:|---:|
| Excel (`.xlsx`) | 4,585 | 2,691 | -41% |
| PDF | 12,333 | 10,306 | -16% |
| PDF, con 10 imágenes descriptas | 12,333 | 11,228 | -9% |
| Word (`.docx`) | 6,004 | 5,999 | -0% |
| PowerPoint (`.pptx`, con notas) | 10,646 | 10,646 | +0% |

El ahorro depende del formato: Excel y los PDF con encabezados repetidos son los que más ganan. Word y PowerPoint casi no cambian, porque MarkItDown ya los convierte bien. En PowerPoint el único ahorro posible es descartar las notas del orador (`--skip-notes`), a costa de perder esa información.

Archivos de prueba propios (un trabajo práctico y una clase de facultad), tokens contados con `tiktoken` (`o200k_base`). Con tus archivos el resultado puede variar: medilo con `--benchmark`.

## Qué hace con cada formato

| Formato | Resultado |
|---|---|
| PDF | Texto por página con marcadores `<!-- Página N/Total -->`, tablas como tablas Markdown, sin encabezados ni pies repetidos, y las páginas con ecuaciones marcadas. Opcional: descripción de las imágenes. |
| Word (`.docx`) | Texto, listas y tablas. Sin marcadores de página (Word no los guarda). |
| Excel (`.xlsx`) | Una tabla por bloque de datos, con sus valores (no las fórmulas), hasta 200 filas por bloque. |
| PowerPoint (`.pptx`) | Una sección por diapositiva. Opcional: descartar las notas del orador. Las imágenes se describen. |
| HTML / CSV | Conversión estándar de MarkItDown. |
| PNG / JPG | Descripción breve de la imagen (requiere configurar un proveedor de IA). |

## Instalación

Requiere Python 3.10 o superior.

1. Descargá el proyecto (**Code → Download ZIP**) o clonalo: `git clone https://github.com/i3450/MD_Converter.git`
2. Abrí una terminal dentro de la carpeta.
3. Recomendado, creá un entorno virtual:
   - Windows: `python -m venv .venv` y luego `.venv\Scripts\activate`
   - Mac / Linux: `python3 -m venv .venv` y luego `source .venv/bin/activate`
4. Instalá las dependencias: `pip install -r requirements.txt`

En Windows, si `python` no funciona, probá con `py`.

## Describir imágenes con IA (opcional)

Sin configurar nada, el script convierte todo pero **no describe imágenes** (las imágenes sueltas se saltean). Para describirlas, necesitás una clave de OpenAI o un recurso de Azure OpenAI:

1. Copiá `.env.example` con el nombre `.env`, en la misma carpeta que `md_converter.py`:
   - Windows: `copy .env.example .env`
   - Mac / Linux: `cp .env.example .env`
2. Abrí `.env` y completá **una** de las dos opciones. Si completás las dos, se usa Azure.

| Proveedor | Variables | Modelo |
|---|---|---|
| OpenAI | `OPENAI_API_KEY` y, opcional, `OPENAI_MODEL` | Por defecto `gpt-5.4-mini` |
| Azure OpenAI | `AZURE_OPENAI_API_KEY`, `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_DEPLOYMENT` | El que hayas desplegado (probado con `gpt-4o`) |

En Azure, `AZURE_OPENAI_DEPLOYMENT` es el nombre que le pusiste al desplegar el modelo, no el nombre del modelo. La clave y el endpoint están en el portal de Azure, en tu recurso de Azure OpenAI.

**Qué modelo elegir.** No hace falta uno grande: la tarea es describir una imagen en 1 a 3 oraciones. Conviene un modelo chico y barato que acepte imágenes, como los de la línea "mini". `gpt-5.4-mini` cuesta, a octubre de 2026, USD 0,75 por millón de tokens de entrada y 4,50 de salida. Los catálogos, regiones y precios cambian seguido: verificalos en la documentación de tu proveedor.

> El soporte de OpenAI directo está implementado pero todavía no se probó con la API real; Azure sí está probado.

**Seguridad de la clave**

- El archivo `.env` está en `.gitignore`: no se sube a GitHub. No lo compartas ni lo pegues en chats.
- Si la clave se filtra, regenerala en el panel de tu proveedor. Borrar el archivo de GitHub no alcanza.
- Configurá un límite o una alerta de gasto.

## Uso

```
python md_converter.py                      # convierte los archivos de la carpeta actual
python md_converter.py -i "C:\ruta\docs"    # otra carpeta de entrada
python md_converter.py -i docs -r           # incluye subcarpetas
python md_converter.py --force              # reconvierte todo
python md_converter.py --benchmark          # compara los tokens contra MarkItDown puro
```

Los resultados quedan en la carpeta `Converted_MD_Files`, conservando la extensión original: `informe.pdf` produce `informe.pdf.md`. Conviene tener los documentos en una carpeta aparte (con `-i`), no mezclados con el código.

Si hay archivos pendientes, el script pregunta (Enter equivale a No):

- **PDF:** ¿agregar una descripción de sus imágenes con IA? Usa la API y suma tokens.
- **PPTX:** ¿descartar las notas del orador? Ahorra tokens, pero se pierden.

### Opciones

| Opción | Qué hace |
|---|---|
| `-i`, `--input` | Carpeta de entrada (por defecto, la actual) |
| `-o`, `--output` | Carpeta de salida (por defecto, `Converted_MD_Files`) |
| `-r`, `--recursive` | Incluye subcarpetas |
| `--force` | Reconvierte aunque ya exista el `.md` |
| `--describe-pdf-images` / `--no-describe-pdf-images` | Describe (o no) las imágenes de los PDF sin preguntar |
| `--skip-notes` / `--keep-notes` | Descarta (o conserva) las notas del orador de los PPTX sin preguntar |
| `--benchmark` | Compara los tokens de cada archivo contra MarkItDown puro y guarda `benchmark.md` |
| `--stats` | Muestra siempre las páginas o secciones más pesadas (por defecto, solo si el archivo supera los 30.000 tokens) |

Si un `.md` ya existe y el original no cambió, se saltea. Si convertiste sin IA y la configurás después, usá `--force` para volver a convertir con las descripciones.

Los tokens se cuentan con `tiktoken`, que viene en `requirements.txt`. Si no puede cargar su vocabulario (la primera vez lo descarga de internet), el script avisa y estima con caracteres ÷ 4.

## Cómo usar los `.md` con una IA

Subí el `.md` al chat. Los archivos con imágenes descriptas empiezan con una nota que le indica a la IA que, antes de dar una cifra que dependa de una imagen, te pida la imagen original. Cuando te diga "página 3, imagen 1", pasale esa imagen o una captura de esa página.

Ejemplo de salida de un PDF:

```
<!-- Informe.pdf | 3 páginas -->

<!-- Página 2/3 -->
...texto de la página...

[Imagen 1: Gráfico de líneas que relaciona el ángulo de apertura con la fuerza del pistón. Valores no transcriptos.]
```

Las páginas con ecuaciones se marcan como `<!-- Página 1/3 | contiene fórmulas -->`, porque las fracciones y los subíndices pueden salir desordenados.

## Costos y privacidad

- Todo se procesa en tu computadora, excepto las imágenes: se envían al proveedor elegido las imágenes sueltas, las de los PPTX y las de los PDF (solo si contestás que sí).
- En los PDF se describen hasta 3 imágenes por página y 30 por archivo, y se ignoran las muy chicas y las repetidas.
- Cada descripción suele ocupar unos 100 tokens.
- Las descripciones de imágenes de PDF se guardan en `.image_cache.json`, dentro de la carpeta de salida: al reconvertir con `--force` no se vuelven a pagar. Si cambia el modelo, se vuelven a pedir.
- Si la API falla 3 veces seguidas en un PDF, el script deja de describir sus imágenes y avisa. Ese `.md` queda sin descripciones: reconvertilo con `--force` cuando lo arregles.

## Limitaciones

- **PDF escaneados:** no hay OCR, salen casi vacíos. Subí el PDF directamente a la IA.
- **Gráficos vectoriales en PDF:** no se describen, solo las imágenes incrustadas.
- **Ecuaciones:** pueden salir desordenadas. Se marcan, pero no se reconstruyen.
- **Excel:** se guardan los valores y no las fórmulas (una fórmula sin valor guardado sale vacía). No se procesan gráficos ni imágenes, y las hojas ocultas se omiten.
- **Word:** sin marcadores de página.
- **Cifras en imágenes:** el script no las transcribe a propósito, porque una lectura errónea pasa desapercibida.

## Problemas frecuentes

| Mensaje o síntoma | Qué revisar |
|---|---|
| `Falta la variable de entorno ...` | El `.env` tiene las variables de Azure a medias, o no está junto a `md_converter.py`. |
| `ModuleNotFoundError` | Corré `pip install -r requirements.txt` con el entorno virtual activado. |
| Error 404 al describir imágenes | En Azure, revisá el endpoint y que `AZURE_OPENAI_DEPLOYMENT` sea el nombre del deployment. En OpenAI, revisá `OPENAI_MODEL`. |
| Error de versión de API en Azure con un modelo nuevo | Probá con una `AZURE_OPENAI_API_VERSION` más reciente (ver la documentación de Azure). |
| `No se pudo describir una imagen ...` | El mensaje trae el error de la API. Revisá que el modelo acepte imágenes y que tengas cuota. |
| Un PDF sale casi vacío | Probablemente esté escaneado. Subilo directo a la IA. |

## Créditos

Basado en [MarkItDown](https://github.com/microsoft/markitdown) (Microsoft). Probado en Windows y Linux con MarkItDown 0.1.x. No probado en macOS.
