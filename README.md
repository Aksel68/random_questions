# 🎲 Randomizador de Pruebas

Genera múltiples formas de una prueba de alternativas reordenando preguntas y opciones de respuesta.

## Requisitos

- Python 3.10+
- Credenciales de Google Drive API (ver `credentials/README.md`)

## Instalación

1. Clonar el repositorio
2. Instalar dependencias con `uv`:
   ```
   uv sync
   ```
3. Configurar credenciales de Google Drive (ver `credentials/README.md`)

## Uso

```
python -m src.main "<ruta_en_drive>" <n_formas>
```

### Ejemplo

```
python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 4
```

Esto genera 4 archivos en la misma carpeta de Drive:
- `Ondas 2026. Ev. 1 F1 Forma A.docx`
- `Ondas 2026. Ev. 1 F1 Forma B.docx`
- `Ondas 2026. Ev. 1 F1 Forma C.docx`
- `Ondas 2026. Ev. 1 F1 Forma D.docx`

## Formato del archivo de entrada

El archivo .docx debe seguir este formato:
- **Preguntas:** Empiezan con `N-` (ej: `1-`, `2-`, `10-`)
- **Alternativas:** Empiezan con `x)` (ej: `a)`, `b)`, `c)`)
- Las preguntas pueden incluir imágenes en el enunciado
- Las alternativas pueden incluir imágenes
- Todo lo que aparece antes de la primera pregunta se considera encabezado y se preserva

## Tests

```
uv run pytest tests/ -v
```
