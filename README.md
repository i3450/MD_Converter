# converter.py

Convierte documentos (PDF, Word, Excel, PowerPoint, HTML, CSV e imágenes) a Markdown para pasárselos a una IA gastando menos tokens y obteniendo mejores resultados.

Está basado en [MarkItDown](https://github.com/microsoft/markitdown) de Microsoft y le agrega:

- Tablas de PDF con las celdas bien separadas (sin números pegados).
- Marcadores de página en los PDF y sin encabezados ni pies de página repetidos.
- Excel dividido en tablas (bloques), sin celdas "NaN" ni redondeos de valores.
- Descripción breve de las imágenes con IA (opcional), pensada para que la IA que lea el `.md` entienda el contexto y te pida la imagen original cuando necesite una cifra.

## Qué hace con cada formato

| Formato | Resultado |
|---|---|
| PDF | Texto por página con marcadores `<!-- Página N/Total -->`, tablas como tablas Markdown, sin encabezados ni pies repetidos, y las páginas con ecuaciones marcadas. Opcional: descripción de las imágenes. |
| Word (`.docx`) | Texto, listas y tablas. No lleva marcadores de página (Word no los guarda). |
| Excel (`.xlsx`) | Una tabla por bloque de datos, con sus valores (no las fórmulas), hasta 200 filas por bloque. |
| PowerPoint (`.pptx`) | Una sección por diapositiva. Opcional: descartar las notas del orador. Las imágenes se describen. |
| HTML / CSV | Conversión estándar de MarkItDown. |
| PNG / JPG | Descripción breve de la imagen (requiere Azure OpenAI). |

## Requisitos

- Python 3.10 o superior.
- Opcional, para describir imágenes: un recurso de Azure OpenAI con un modelo que acepte imágenes (por ejemplo `gpt-4o`). Sin esto el script funciona igual, pero no describe imágenes.

## Instalación

1. Descargá el proyecto (botón **Code → Download ZIP**) o clonalo:
   ```
   git clone https://github.com/<tu-usuario>/<tu-repo>.git
   ```
2. Abrí una terminal dentro de la carpeta del proyecto.
3. Recomendado, creá un entorno virtual:
   - Windows: `python -m venv .venv` y luego `.venv\Scripts\activate`
   - Mac / Linux: `python3 -m venv .venv` y luego `source .venv/bin/activate`
4. Instalá las dependencias:
   ```
   pip install -r requirements.txt
   ```

En Windows, si `python` no funciona, probá con `py`.

## Configuración (opcional)

Sin configurar nada, el script convierte todo pero **no describe imágenes** (las imágenes sueltas se saltean).

Para describir imágenes con Azure OpenAI:

1. Copiá `.env.example` con el nombre `.env`, en la misma carpeta que `converter.py`:
   - Windows: `copy .env.example .env`
   - Mac / Linux: `cp .env.example .env`
2. Abrí `.env` y completá los tres valores:
   - `AZURE_OPENAI_API_KEY`: la clave de tu recurso.
   - `AZURE_OPENAI_ENDPOINT`: la URL de tu recurso.
   - `AZURE_OPENAI_DEPLOYMENT`: el nombre que le pusiste al desplegar el modelo (no el nombre del modelo).

   Clave y endpoint están en el portal de Azure, en tu recurso de Azure OpenAI, en la sección de claves y punto de conexión. Los nombres de los menús pueden cambiar.

**Seguridad de la clave**

- El archivo `.env` está en `.gitignore`: no se sube a GitHub. No lo compartas ni lo pegues en chats.
- Si la clave se filtra, regenerala en el portal de Azure. Borrar el archivo de GitHub no alcanza.
- Configurá una alerta o un límite de gasto en Azure.

## Uso

```
python converter.py                      # convierte los archivos de la carpeta actual
python converter.py -i "C:\ruta\docs"    # otra carpeta de entrada
python converter.py -i docs -r           # incluye subcarpetas
python converter.py --force              # reconvierte todo
```

Los resultados quedan en la carpeta `Converted_MD_Files`, conservando la extensión original: `informe.pdf` produce `informe.pdf.md`.

Conviene tener los documentos en una carpeta aparte (con `-i`) y no mezclados con el código.

### Preguntas que hace el script

Si hay archivos pendientes de convertir, pregunta (Enter equivale a No):

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

Si un `.md` ya existe y el original no cambió, se saltea. Si convertiste sin Azure y lo configurás después, usá `--force` para que los PDF y PPTX con imágenes se vuelvan a convertir con sus descripciones.

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

- Todo se procesa en tu computadora, excepto las imágenes: se envían a Azure las imágenes sueltas, las de los PPTX y las de los PDF (solo si contestás que sí).
- En los PDF se describen hasta 3 imágenes por página y 30 por archivo, y se ignoran las muy chicas y las repetidas.
- Cada descripción suele ocupar unos 100 tokens.
- Los tokens que muestra el script son estimados (caracteres ÷ 4).

## Limitaciones

- **PDF escaneados:** no hay OCR, salen casi vacíos. Subí el PDF directamente a la IA.
- **Gráficos vectoriales en PDF:** no se describen, solo las imágenes incrustadas.
- **Ecuaciones:** pueden salir desordenadas. Se marcan, pero no se reconstruyen.
- **Excel:** se guardan los valores y no las fórmulas. No se procesan gráficos ni imágenes, y las hojas ocultas se omiten.
- **Word:** sin marcadores de página.
- **Cifras en imágenes:** el script no las transcribe a propósito, porque una lectura errónea pasa desapercibida.

## Problemas frecuentes

| Mensaje o síntoma | Qué revisar |
|---|---|
| `Falta la variable de entorno ...` | El `.env` está incompleto o no está junto a `converter.py`. |
| `ModuleNotFoundError` | Corré `pip install -r requirements.txt` con el entorno virtual activado. |
| Error 404 al describir imágenes | Revisá el endpoint y que `AZURE_OPENAI_DEPLOYMENT` sea el nombre del deployment. |
| `No se pudo describir una imagen ...` | El mensaje trae el error de la API. Revisá que el modelo acepte imágenes y que tengas cuota. |
| Un PDF sale casi vacío | Probablemente esté escaneado. Subilo directo a la IA. |

## Créditos

Basado en [MarkItDown](https://github.com/microsoft/markitdown) (Microsoft). Probado en Windows y Linux con MarkItDown 0.1.x. No probado en macOS.
