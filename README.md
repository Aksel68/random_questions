# 🎲 Randomizador de Pruebas

Genera múltiples formas de una prueba de alternativas a partir de un archivo `.docx` almacenado en Google Drive. Reordena preguntas y alternativas aleatoriamente, preserva todo el formato original (imágenes, estilos, tablas), y genera automáticamente las pautas de respuestas en `.xlsx`.

---

## ¿Cómo funciona?

1. **Descarga** el `.docx` de la forma original desde Google Drive.
2. **Parsea** las preguntas y alternativas, preservando el XML completo de cada párrafo (formato, imágenes, estilos).
3. **Genera N formas** nuevas reordenando preguntas y alternativas aleatoriamente.
4. **Construye** un `.docx` por forma clonando el original y reemplazando solo el contenido de la tabla de preguntas.
5. **Descarga la pauta** original desde un Google Sheet en la misma carpeta de Drive y genera una pauta `.xlsx` por forma con las respuestas correctas remapeadas.
6. **Sube** todas las formas y pautas a la misma carpeta de Drive.

---

## Estructura del proyecto

```
random_questions/
├── src/
│   ├── main.py          # Punto de entrada — orquesta todo el flujo
│   ├── models.py        # Dataclasses: Exam, Question, Alternative
│   ├── parser.py        # Lee el .docx y extrae preguntas con XML preservado
│   ├── randomizer.py    # Genera N formas con preguntas/alternativas shuffleadas
│   ├── doc_builder.py   # Construye el .docx de cada forma clonando el original
│   ├── answer_key.py    # Parsea la pauta original y genera pautas por forma (.xlsx)
│   └── drive_client.py  # Descarga y sube archivos desde/hacia Google Drive
├── tests/               # Tests unitarios e integración con pytest
├── docs/plans/          # Planes de desarrollo y decisiones de diseño
├── credentials/
│   └── README.md        # Instrucciones para configurar Google Drive API
├── debug_q5.py          # Script de diagnóstico para inspeccionar párrafos de una pregunta
├── test_q5_fix.py       # Script de verificación del fix de imágenes en alternativas
├── pyproject.toml
└── README.md
```

---

## Requisitos

- Python 3.10+
- [`uv`](https://github.com/astral-sh/uv) (gestor de paquetes)
- Credenciales de Google Drive API (ver `credentials/README.md`)

### Dependencias principales

| Paquete | Uso |
|---|---|
| `python-docx` | Leer y escribir archivos `.docx` |
| `lxml` | Manipulación directa del XML de Word |
| `openpyxl` | Generar pautas de respuestas en `.xlsx` |
| `google-api-python-client` | Interactuar con Google Drive API |
| `google-auth-oauthlib` | Autenticación OAuth2 con Google |

---

## Instalación

```bash
# 1. Clonar el repositorio
git clone https://github.com/Aksel68/random_questions.git
cd random_questions

# 2. Instalar dependencias
uv sync

# 3. Configurar credenciales de Google Drive
# → Seguir las instrucciones en credentials/README.md
```

---

## Configuración de Google Drive API

1. Ir a [Google Cloud Console](https://console.cloud.google.com/) y crear un proyecto.
2. Habilitar la **Google Drive API** en *APIs & Services > Library*.
3. Crear credenciales OAuth 2.0 (*Desktop app*) y descargar el JSON.
4. Guardarlo como `credentials/client_secret.json`.
5. En la primera ejecución se abrirá el navegador para autenticarse. Esto genera `credentials/token.json` automáticamente.

> ⚠️ **Nunca subas** `client_secret.json` ni `token.json` al repositorio. Están en `.gitignore`.

---

## Uso

```bash
uv run python -m src.main "<ruta_en_drive>" <n_formas>
```

### Ejemplo

```bash
uv run python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 4
```

Esto genera en la misma carpeta de Drive:

| Archivo | Descripción |
|---|---|
| `Ondas 2026. Ev. 1 F2.docx` | Forma 2 — preguntas y alternativas reordenadas |
| `Ondas 2026. Ev. 1 F3.docx` | Forma 3 |
| `Ondas 2026. Ev. 1 F4.docx` | Forma 4 |
| `Ondas 2026. Ev. 1 F5.docx` | Forma 5 |
| `Claves Ondas 2026. Ev. 1 F2.xlsx` | Pauta de respuestas forma 2 |
| `Claves Ondas 2026. Ev. 1 F3.xlsx` | Pauta de respuestas forma 3 |
| `Claves Ondas 2026. Ev. 1 F4.xlsx` | Pauta de respuestas forma 4 |
| `Claves Ondas 2026. Ev. 1 F5.xlsx` | Pauta de respuestas forma 5 |

---

## Formato del archivo de entrada

### Nombre del archivo
Debe terminar en `F<número>.docx`, por ejemplo:
- `Prueba Física F1.docx` → genera F2, F3, ...
- `Evaluación 2 F3.docx` → genera F4, F5, ...

### Estructura del documento
- **Preguntas:** El enunciado empieza con `N-` o `N.` (ej: `1-`, `2.`, `10-`)
- **Alternativas:** Empiezan con `x)` (ej: `a)`, `b)`, `c)`, `d)`, `e)`)
- **Imágenes:** Pueden estar en el enunciado o en las alternativas — se reposicionan automáticamente al enunciado al construir la forma
- **Forma:** El documento debe contener el texto `Forma N` (se actualiza automáticamente)
- **Tabla de preguntas:** Las preguntas deben estar en una tabla (puede ser de 1 o 2 columnas)

### Pauta de respuestas (opcional)
Se busca automáticamente en la misma carpeta de Drive un Google Sheet llamado `Claves <nombre_archivo_sin_extensión>`.

Ejemplo: para `Ondas F1.docx` se busca `Claves Ondas F1`.

Formato esperado del sheet:

| Pregunta | Respuesta |
|---|---|
| 1 | a |
| 2 | c |
| ... | ... |

Si no se encuentra la pauta, se generan igual las formas `.docx` sin pautas `.xlsx`.

---

## Tests

```bash
uv run pytest tests/ -v
```

---

## Scripts de diagnóstico

### `debug_q5.py` — Inspeccionar estructura de párrafos

Muestra cómo el parser clasifica cada párrafo de una pregunta específica, útil para diagnosticar problemas de posicionamiento de imágenes.

```bash
# Desde Drive
uv run python debug_q5.py --drive "Guías/Ondas 2026. Ev. 1 F2.docx" --question 5

# Desde archivo local
uv run python debug_q5.py --local /ruta/al/archivo.docx --question 3
```

### `test_q5_fix.py` — Verificar posicionamiento de imágenes

Verifica que las imágenes embebidas en alternativas se extraigan correctamente al enunciado.

```bash
uv run python test_q5_fix.py --drive "Guías/Ondas 2026. Ev. 1 F2.docx"
```
