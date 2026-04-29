# Fix Page Breaks & Form Naming — Plan de Implementación

**Fecha:** 2026-04-27
**Contexto:** Tres correcciones necesarias al flujo actual del randomizador.

## Problemas identificados

### Problema 1: Preguntas cortadas por salto de página
Cuando el documento generado tiene preguntas (celdas de tabla) que no caben en
el espacio restante de la página actual, Word las corta a la mitad. Es mejor
forzar que la fila de la tabla comience en la página siguiente si no cabe completa.

### Problema 2: Nomenclatura de archivos generados
El archivo original ya termina en `F1` (e.g. `Ondas 2026. Ev. 1 F1.docx`).
Las formas generadas deben nombrarse `F2`, `F3`, `F4`... en lugar de
`Forma A`, `Forma B`, `Forma C`.

**Antes:**
```
Ondas 2026. Ev. 1 F1.docx          → entrada
Ondas 2026. Ev. 1 F1 Forma A.docx  → generado (incorrecto)
Ondas 2026. Ev. 1 F1 Forma B.docx  → generado (incorrecto)
```

**Después:**
```
Ondas 2026. Ev. 1 F1.docx          → entrada
Ondas 2026. Ev. 1 F2.docx          → generado (correcto)
Ondas 2026. Ev. 1 F3.docx          → generado (correcto)
```

### Problema 3: Texto "Forma X" dentro del documento
El documento original contiene un texto debajo del título que dice "Forma 1".
Los documentos generados deben actualizar ese texto a "Forma 2", "Forma 3", etc.
según el número de forma correspondiente al archivo.

---

## Estructura de archivos a modificar

```
src/
├── main.py           # Cambiar nomenclatura de archivos y sufijo de forma
├── doc_builder.py    # Agregar page-break awareness + adaptar label "Forma N"
└── models.py         # (sin cambios)
tests/
├── test_doc_builder.py  # Tests para page breaks y label numérico
└── test_integration.py  # Actualizar para nueva nomenclatura
docs/
└── plans/
    └── 2026-04-27-fix-page-breaks-and-form-naming.md  # Este archivo
```

---

## Tarea 1: Cambiar nomenclatura de archivos — `main.py`

### Análisis

La función `get_form_suffix` actualmente genera letras (A, B, C...).
Necesitamos:
1. Detectar el número de forma del archivo original (extraer `N` de `FN`).
2. Generar nombres de archivo reemplazando `FN` por `F(N+1)`, `F(N+2)`, etc.
3. Pasar el número de forma (no la letra) al `doc_builder`.

### Cambios en `main.py`

```python
import re

def extract_form_number(filename: str) -> tuple[str, int]:
    """Extrae el número de forma y el prefijo base del nombre de archivo.

    Busca el patrón 'F<número>' al final del nombre base (antes de .docx).

    Args:
        filename: Nombre del archivo (e.g. "Ondas 2026. Ev. 1 F1.docx").

    Returns:
        Tupla (base_sin_FN, número_de_forma).
        Ejemplo: ("Ondas 2026. Ev. 1 ", 1)

    Raises:
        ValueError: Si no se encuentra el patrón F<número>.
    """
    base_name = filename.rsplit(".docx", 1)[0]
    match = re.search(r"^(.*\bF)(\d+)$", base_name)
    if not match:
        raise ValueError(
            f"No se encontró el patrón 'F<número>' en el nombre del archivo: {filename}. "
            f"El archivo debe terminar en F1, F2, etc. antes de la extensión .docx"
        )
    prefix = match.group(1)       # "Ondas 2026. Ev. 1 F"
    form_number = int(match.group(2))  # 1
    return prefix, form_number
```

Luego, en `main()`, reemplazar la lógica de generación de nombres:

```python
def main():
    # ... (args parsing sin cambios) ...

    path_parts = drive_path.strip("/").split("/")
    filename = path_parts[-1]
    drive_folder = "/".join(path_parts[:-1]) if len(path_parts) > 1 else ""

    # Extraer información del nombre del archivo
    prefix, original_form_number = extract_form_number(filename)

    print(f"📄 Archivo: {drive_path}")
    print(f"📁 Carpeta destino: {drive_folder or 'raíz de Drive'}")
    print(f"🔢 Forma original: F{original_form_number}")
    print(f"🔄 Generando {n_forms} forma(s)...\n")

    # ... (descarga, parseo, generación sin cambios) ...

    # 4. Construir y subir cada forma
    temp_dir = tempfile.mkdtemp()
    for i, form_questions in enumerate(forms):
        new_form_number = original_form_number + 1 + i  # F1 → F2, F3, F4...
        form_filename = f"{prefix}{new_form_number}.docx"
        form_path = os.path.join(temp_dir, form_filename)

        print(f"\n📝 Generando: {form_filename}")
        build_docx(
            source_path=local_path,
            questions=form_questions,
            output_path=form_path,
            form_number=new_form_number,  # Ahora pasamos número, no letra
            questions_table_index=exam.questions_table_index,
            questions_per_row=exam.questions_per_row,
        )

        print(f"⬆️  Subiendo a Drive...")
        upload_docx(form_path, drive_folder, form_filename)

        os.unlink(form_path)

    # ... (cleanup sin cambios) ...
```

### Eliminar `get_form_suffix`

Ya no se necesita. Se elimina la función completa.

### Tests — `tests/test_main.py` (nuevo archivo)

```python
import pytest
from src.main import extract_form_number


def test_extract_form_number_simple():
    prefix, num = extract_form_number("Ondas 2026. Ev. 1 F1.docx")
    assert prefix == "Ondas 2026. Ev. 1 F"
    assert num == 1


def test_extract_form_number_higher():
    prefix, num = extract_form_number("Prueba F5.docx")
    assert prefix == "Prueba F"
    assert num == 5


def test_extract_form_number_no_match():
    with pytest.raises(ValueError, match="No se encontró"):
        extract_form_number("Prueba sin forma.docx")


def test_extract_form_number_multi_digit():
    prefix, num = extract_form_number("Test F12.docx")
    assert prefix == "Test F"
    assert num == 12
```

---

## Tarea 2: Actualizar label "Forma X" dentro del documento — `doc_builder.py`

### Análisis

La función `_update_form_label` actualmente recibe `form_suffix: str` (la letra)
y reemplaza "Forma X" → "Forma A". Ahora debe recibir un número entero y
reemplazar "Forma 1" → "Forma 2", etc.

### Cambios en `_update_form_label`

```python
def _update_form_label(doc: Document, form_number: int):
    """Busca y reemplaza el texto 'Forma N' en los párrafos superiores.

    Busca un párrafo que contenga la palabra "Forma" seguida de un número
    y lo reemplaza con el nuevo número de forma.

    Args:
        doc: Documento python-docx.
        form_number: Número de la forma (ej: 2, 3, 4).
    """
    form_pattern = re.compile(r"([Ff]orma\s+)\d+")

    for paragraph in doc.paragraphs:
        full_text = paragraph.text
        if form_pattern.search(full_text):
            for run in paragraph.runs:
                if form_pattern.search(run.text):
                    run.text = form_pattern.sub(
                        rf"\g<1>{form_number}", run.text
                    )
            break  # Solo reemplazar la primera ocurrencia
```

### Cambios en la firma de `build_docx`

```python
def build_docx(
    source_path: str,
    questions: list[Question],
    output_path: str,
    form_number: int,               # Antes: form_suffix: str
    questions_table_index: int = 2,
    questions_per_row: int = 2,
) -> None:
    # ... (copiar archivo) ...

    # 3. Actualizar label "Forma X"
    _update_form_label(doc, form_number)  # Antes: _update_form_label(doc, form_suffix)

    # ... (resto sin cambios) ...
```

### Tests — actualizar `tests/test_doc_builder.py`

Todos los tests que pasan `form_suffix="A"` deben cambiarse a `form_number=2`.
Los asserts que verifican "Forma A" deben verificar "Forma 2".

---

## Tarea 3: Page-break awareness en las filas de la tabla — `doc_builder.py`

### Análisis

Word permite configurar propiedades de fila en tablas para evitar que se
corten entre páginas. Hay DOS mecanismos complementarios:

1. **`<w:cantSplit>`** en `<w:trPr>`: Indica que la fila NO debe dividirse
   entre páginas. Si la fila no cabe en la página actual, Word la mueve
   completa a la siguiente.

2. **`<w:tblHeader>`** en `<w:trPr>`: Marca filas como encabezado de tabla
   (se repiten en cada página). No aplica aquí.

El mecanismo correcto es aplicar `<w:cantSplit/>` a cada fila de la tabla
de preguntas. Esto le dice a Word: "si esta fila no cabe completa en lo
que queda de página, ponla al inicio de la siguiente".

### Implementación en `doc_builder.py`

```python
from lxml.etree import SubElement

def _set_rows_cant_split(table):
    """Configura todas las filas de la tabla para que no se dividan entre páginas.

    Agrega la propiedad <w:cantSplit/> a <w:trPr> de cada fila.
    Si la fila no tiene <w:trPr>, lo crea.

    Esto asegura que si una pregunta no cabe en el espacio restante de la
    página actual, Word mueve la fila completa al inicio de la siguiente página.

    Args:
        table: Objeto Table de python-docx.
    """
    nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    for row in table.rows:
        tr = row._tr
        # Buscar o crear <w:trPr>
        trPr = tr.find(f"{{{w_ns}}}trPr")
        if trPr is None:
            trPr = SubElement(tr, f"{{{w_ns}}}trPr")
            # Insertar <w:trPr> como primer hijo de <w:tr>
            tr.remove(trPr)
            tr.insert(0, trPr)

        # Buscar o crear <w:cantSplit>
        cant_split = trPr.find(f"{{{w_ns}}}cantSplit")
        if cant_split is None:
            SubElement(trPr, f"{{{w_ns}}}cantSplit")
```

### Integración en `build_docx`

Agregar la llamada después de obtener la tabla de preguntas y antes de
reemplazar el contenido de las celdas:

```python
def build_docx(
    source_path: str,
    questions: list[Question],
    output_path: str,
    form_number: int,
    questions_table_index: int = 2,
    questions_per_row: int = 2,
) -> None:
    # 1. Copiar el archivo original
    shutil.copy2(source_path, output_path)

    # 2. Abrir la copia y modificar
    doc = Document(output_path)

    # 3. Actualizar label "Forma X"
    _update_form_label(doc, form_number)

    # 4. Obtener la tabla de preguntas
    questions_table = doc.tables[questions_table_index]

    # 5. Configurar filas para que no se corten entre páginas
    _set_rows_cant_split(questions_table)

    # 6. Reemplazar contenido de cada celda
    q_index = 0
    for row in questions_table.rows:
        for col_idx in range(questions_per_row):
            if q_index >= len(questions):
                break
            cell = row.cells[col_idx]
            _clear_cell(cell)
            cell_content = _build_cell_content(questions[q_index])
            _add_paragraphs_to_cell(cell, cell_content)
            q_index += 1

    # 7. Guardar
    doc.save(output_path)
```

### Tests — `tests/test_doc_builder.py`

```python
def test_build_docx_sets_cant_split_on_rows(tmp_path):
    """Verifica que las filas de la tabla de preguntas tienen cantSplit."""
    source = _create_source_exam(tmp_path)
    output = tmp_path / "output.docx"
    questions = [...]  # crear preguntas de test

    build_docx(
        source_path=str(source),
        questions=questions,
        output_path=str(output),
        form_number=2,
        questions_table_index=2,
        questions_per_row=2,
    )

    doc = Document(str(output))
    questions_table = doc.tables[2]
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    for row in questions_table.rows:
        tr = row._tr
        trPr = tr.find(f"{{{w_ns}}}trPr")
        assert trPr is not None, "La fila debe tener <w:trPr>"
        cant_split = trPr.find(f"{{{w_ns}}}cantSplit")
        assert cant_split is not None, "La fila debe tener <w:cantSplit>"
```

---

## Tarea 4: Actualizar README y documentación

### Cambios en `README.md`

Actualizar el ejemplo de uso y los nombres de archivos generados:

```markdown
### Ejemplo

    uv run python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 4

Esto genera 4 archivos en la misma carpeta de Drive:
- `Ondas 2026. Ev. 1 F2.docx`
- `Ondas 2026. Ev. 1 F3.docx`
- `Ondas 2026. Ev. 1 F4.docx`
- `Ondas 2026. Ev. 1 F5.docx`

Cada archivo generado contiene el texto "Forma N" actualizado
según su número de forma.
```

### Actualizar descripción de formato

Agregar nota sobre el requisito de nomenclatura:

```markdown
## Formato del archivo de entrada

- El nombre del archivo **debe** terminar en `F<número>.docx` (ej: `F1.docx`, `F2.docx`)
- El documento debe contener un texto "Forma N" debajo del título
```

---

## Tarea 5: Actualizar test de integración — `tests/test_integration.py`

Los tests de integración que verifican nombres de archivo y labels de forma
deben actualizarse para usar la nueva nomenclatura numérica.

### Cambios clave

- Verificar que los archivos generados se llamen `...F2.docx`, `...F3.docx`
- Verificar que el texto dentro del documento diga "Forma 2", "Forma 3"
- Verificar que las filas de la tabla tengan `<w:cantSplit>`

---

## Resumen de tareas

| #   | Tarea                                         | Archivos                                   | Dependencias |
| --- | --------------------------------------------- | ------------------------------------------ | ------------ |
| 1   | Cambiar nomenclatura de archivos              | `src/main.py`, `tests/test_main.py` (new)  | —            |
| 2   | Actualizar label "Forma X" a "Forma N"        | `src/doc_builder.py`, `tests/test_doc_builder.py` | 1      |
| 3   | Page-break awareness (cantSplit)              | `src/doc_builder.py`, `tests/test_doc_builder.py` | —      |
| 4   | Actualizar README y documentación             | `README.md`                                | 1, 2         |
| 5   | Actualizar tests de integración               | `tests/test_integration.py`                | 1, 2, 3      |

## Decisiones de diseño clave

1. **`<w:cantSplit>` vs saltos de página manuales:** Se eligió `cantSplit` porque
   es la forma estándar de Word para evitar que filas se dividan entre páginas.
   Los saltos de página manuales (`<w:pageBreakBefore>`) serían frágiles y
   requerirían calcular alturas de contenido, lo cual python-docx no soporta.

2. **Numeración secuencial vs letras:** El cambio de `Forma A/B/C` a `F2/F3/F4`
   refleja la convención real del usuario. El número de forma en el nombre del
   archivo y el número dentro del texto "Forma N" deben coincidir.

3. **Regex para detectar `FN`:** Se usa `\bF(\d+)$` al final del nombre base
   para evitar falsos positivos con otras ocurrencias de "F" en el nombre.
   El `\b` (word boundary) evita que matchee con palabras como "STAFF1".