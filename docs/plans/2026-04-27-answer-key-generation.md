# Generación de Pautas de Respuestas — Plan de Implementación

> **Para agentes:** Usá `@rules/skills/subagents.md` (recomendado) o `@rules/skills/executing-plans.md` para implementar este plan tarea por tarea. Los pasos usan sintaxis checkbox (`- [ ]`) para tracking.

**Goal:** Generar automáticamente las pautas (claves de respuesta) de cada forma randomizada, a partir de la pauta original almacenada como Google Sheet en Drive.

**Arquitectura:** La pauta original mapea `{número_pregunta → letra_correcta}`. El randomizer mezcla tanto el orden de preguntas como el orden de alternativas. Para generar la pauta de una forma nueva necesitamos trackear: (a) el número original de cada pregunta antes de renumerar, y (b) la letra original de cada alternativa antes de reasignar. Con esos dos datos, el mapeo es directo: buscar la respuesta correcta original → encontrar a qué letra nueva corresponde → escribir en la pauta de la nueva forma.

**Tech Stack:** openpyxl (nueva dependencia) para escribir .xlsx, Google Drive API para descargar la pauta original (Google Sheets nativo exportado como CSV) y subir las pautas generadas.

---

## Flujo de datos

```
Pauta original (Google Sheet):          Forma randomizada:
  Q1 → a (correcta)                      Q15 (era Q1 original):
  Q2 → b (correcta)                        a) [era d)] ← NO
  ...                                       b) [era b)] ← NO
                                            c) [era a)] ← SÍ ← nueva respuesta correcta
                                            d) [era c)] ← NO

Pauta nueva (F2):
  Q15 → c
```

## Estructura de archivos

```
src/
├── models.py           # Modificar: agregar original_number y original_letter
├── randomizer.py       # Modificar: preservar originales durante shuffle/renumber
├── answer_key.py       # Crear: parsear pauta original + generar pautas nuevas
├── drive_client.py     # Modificar: agregar download_sheet y upload_xlsx
├── main.py             # Modificar: orquestar descarga de pauta + generación + subida
└── doc_builder.py      # Sin cambios
tests/
├── test_answer_key.py  # Crear: tests para parseo y generación de pautas
├── test_models.py      # Modificar: tests para original_number/original_letter
└── test_randomizer.py  # Modificar: tests para preservación de originales
pyproject.toml          # Modificar: agregar openpyxl
```

---

### Tarea 1: Agregar campos de rastreo a los modelos

**Archivos:**
- Modificar: `src/models.py`
- Test: `tests/test_models.py`

El randomizer actualmente sobrescribe `Alternative.letter` y `Question.number` durante el shuffle. Necesitamos campos que preserven los valores originales para poder reconstruir la pauta.

- [ ] **Paso 1: Escribir tests para los nuevos campos**

```python
# En tests/test_models.py — agregar estos tests

def test_alternative_original_letter_defaults_to_letter():
    """original_letter debe ser igual a letter si no se especifica."""
    from lxml import etree
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    p = etree.SubElement(etree.Element("root"), f"{{{ns}}}p")
    r = etree.SubElement(p, f"{{{ns}}}r")
    t = etree.SubElement(r, f"{{{ns}}}t")
    t.text = "a) Opción A"

    alt = Alternative(letter="a", paragraphs=[p])
    assert alt.original_letter == "a"


def test_alternative_original_letter_preserved_on_clone():
    """clone() debe preservar original_letter."""
    from lxml import etree
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    p = etree.SubElement(etree.Element("root"), f"{{{ns}}}p")

    alt = Alternative(letter="c", paragraphs=[p], original_letter="a")
    cloned = alt.clone()
    assert cloned.original_letter == "a"
    assert cloned.letter == "c"


def test_question_original_number_defaults_to_number():
    """original_number debe ser igual a number si no se especifica."""
    q = Question(number=5, enunciado_paragraphs=[])
    assert q.original_number == 5


def test_question_original_number_preserved_on_clone():
    """clone() debe preservar original_number."""
    q = Question(number=12, enunciado_paragraphs=[], original_number=3)
    cloned = q.clone()
    assert cloned.original_number == 3
    assert cloned.number == 12
```

- [ ] **Paso 2: Correr tests para verificar que fallan**

Comando: `uv run pytest tests/test_models.py -v`
Esperado: FAIL — `Alternative.__init__` y `Question.__init__` no aceptan los nuevos campos.

- [ ] **Paso 3: Agregar `original_letter` a `Alternative`**

En `src/models.py`, modificar la dataclass `Alternative`:

```python
@dataclass
class Alternative:
    letter: str
    paragraphs: list  # list of lxml elements (<w:p>)
    original_letter: str = ""

    def __post_init__(self):
        if not self.original_letter:
            self.original_letter = self.letter

    def clone(self) -> "Alternative":
        return Alternative(
            letter=self.letter,
            paragraphs=[deepcopy(p) for p in self.paragraphs],
            original_letter=self.original_letter,
        )
```

- [ ] **Paso 4: Agregar `original_number` a `Question`**

En `src/models.py`, modificar la dataclass `Question`:

```python
@dataclass
class Question:
    number: int
    enunciado_paragraphs: list  # list of lxml elements (<w:p>)
    alternatives: list[Alternative] = field(default_factory=list)
    trailing_empty_paragraphs: list = field(default_factory=list)
    original_number: int = 0

    def __post_init__(self):
        if self.original_number == 0:
            self.original_number = self.number

    def clone(self) -> "Question":
        return Question(
            number=self.number,
            enunciado_paragraphs=[deepcopy(p) for p in self.enunciado_paragraphs],
            alternatives=[alt.clone() for alt in self.alternatives],
            trailing_empty_paragraphs=[
                deepcopy(p) for p in self.trailing_empty_paragraphs
            ],
            original_number=self.original_number,
        )
```

- [ ] **Paso 5: Correr todos los tests**

Comando: `uv run pytest tests/ -v`
Esperado: PASS — los nuevos campos tienen defaults retrocompatibles, nada se rompe.

- [ ] **Paso 6: Commit**

```bash
git add src/models.py tests/test_models.py
git commit -m "feat: add original_number/original_letter tracking to models"
```

---

### Tarea 2: Preservar originales en el randomizer

**Archivos:**
- Modificar: `src/randomizer.py`
- Test: `tests/test_randomizer.py`

Actualmente `_renumber_question` sobrescribe `question.number` y `_relabel_alternative` sobrescribe `alt.letter`. Después del clone + shuffle, los valores originales se pierden. Ahora que los modelos tienen `original_number` y `original_letter`, el clone ya los preserva. Solo necesitamos asegurar que `_renumber_question` y `_relabel_alternative` NO sobrescriban los campos originales, y escribir tests que verifiquen la preservación end-to-end.

- [ ] **Paso 1: Escribir tests de preservación**

```python
# En tests/test_randomizer.py — agregar estos tests

def test_generate_forms_preserves_original_numbers():
    """Después de randomizar, cada pregunta conserva original_number."""
    exam = _create_test_exam()  # helper existente
    forms = generate_forms(exam, n=1)

    original_numbers = {q.number for q in exam.questions}
    form_original_numbers = {q.original_number for q in forms[0]}
    assert form_original_numbers == original_numbers, (
        "Los original_number de la forma deben coincidir con los numbers del exam original"
    )


def test_generate_forms_preserves_original_letters():
    """Después de randomizar, cada alternativa conserva original_letter."""
    exam = _create_test_exam()
    forms = generate_forms(exam, n=1)

    for q in forms[0]:
        original_letters = {alt.original_letter for alt in q.alternatives}
        # Debe tener las 4 letras originales (a, b, c, d)
        assert original_letters == {"a", "b", "c", "d"}, (
            f"Q{q.number} (original Q{q.original_number}): "
            f"original_letters={original_letters}, esperado={{'a','b','c','d'}}"
        )


def test_generate_forms_original_vs_current_differ():
    """En al menos alguna pregunta, letter != original_letter (prueba de shuffle real)."""
    import random
    random.seed(42)
    exam = _create_test_exam()
    forms = generate_forms(exam, n=1)

    found_relabeled = False
    for q in forms[0]:
        for alt in q.alternatives:
            if alt.letter != alt.original_letter:
                found_relabeled = True
                break
        if found_relabeled:
            break
    assert found_relabeled, (
        "Con seed=42 debería haber al menos una alternativa relabeled"
    )
```

- [ ] **Paso 2: Correr tests para verificar que fallan**

Comando: `uv run pytest tests/test_randomizer.py -v -k "original"`
Esperado: FAIL — `_renumber_question` sobrescribe `question.number` pero `original_number` ya se preserva vía clone. Sin embargo, `_relabel_alternative` sobrescribe `alt.letter` sin preservar `original_letter`. Verificar exactamente qué falla.

- [ ] **Paso 3: Verificar que `_renumber_question` NO toca `original_number`**

Revisar `_renumber_question` en `src/randomizer.py`. Actualmente hace:
```python
question.number = new_number
```

Esto está bien — `original_number` se setea en el clone (desde `__post_init__`) y `_renumber_question` solo toca `.number`. No requiere cambios.

- [ ] **Paso 4: Verificar que `_relabel_alternative` NO toca `original_letter`**

Revisar `_relabel_alternative` en `src/randomizer.py`. Actualmente hace:
```python
alt.letter = new_letter
```

Esto está bien — `original_letter` se setea en el clone y `_relabel_alternative` solo toca `.letter`. No requiere cambios.

- [ ] **Paso 5: Correr tests**

Comando: `uv run pytest tests/test_randomizer.py -v`
Esperado: PASS — los campos originales se preservan automáticamente gracias a los defaults en `__post_init__` y el clone.

- [ ] **Paso 6: Correr suite completa**

Comando: `uv run pytest tests/ -v`
Esperado: PASS — nada roto.

- [ ] **Paso 7: Commit**

```bash
git add tests/test_randomizer.py
git commit -m "test: verify original_number/original_letter preservation in randomizer"
```

---

### Tarea 3: Crear módulo de pauta de respuestas

**Archivos:**
- Crear: `src/answer_key.py`
- Crear: `tests/test_answer_key.py`

Este módulo se encarga de: (1) parsear la pauta original (CSV exportado desde Google Sheets), y (2) generar la pauta de una forma randomizada como .xlsx.

- [ ] **Paso 1: Agregar openpyxl como dependencia**

En `pyproject.toml`, agregar `openpyxl` a dependencies:

```toml
dependencies = [
    "python-docx==1.1.2",
    "google-api-python-client==2.131.0",
    "google-auth-oauthlib==1.2.0",
    "google-auth-httplib2==0.2.0",
    "openpyxl>=3.1.0",
]
```

Correr: `uv sync`

- [ ] **Paso 2: Escribir tests para parseo de pauta original**

```python
# tests/test_answer_key.py

import os
import tempfile

import pytest

from src.answer_key import parse_answer_key_csv, generate_answer_key, build_answer_key_xlsx


class TestParseAnswerKeyCsv:
    """Tests para parseo de la pauta original desde CSV."""

    def test_parses_simple_csv(self):
        csv_content = "Pregunta,Respuesta\n1,a\n2,b\n3,c\n4,d\n"
        result = parse_answer_key_csv(csv_content)
        assert result == {1: "a", 2: "b", 3: "c", 4: "d"}

    def test_parses_36_questions(self):
        lines = ["Pregunta,Respuesta"]
        letters = "abcd"
        for i in range(1, 37):
            lines.append(f"{i},{letters[(i - 1) % 4]}")
        csv_content = "\n".join(lines) + "\n"
        result = parse_answer_key_csv(csv_content)
        assert len(result) == 36
        assert result[1] == "a"
        assert result[36] == "d"

    def test_strips_whitespace(self):
        csv_content = "Pregunta,Respuesta\n 1 , a \n 2 , b \n"
        result = parse_answer_key_csv(csv_content)
        assert result == {1: "a", 2: "b"}

    def test_empty_csv_returns_empty(self):
        csv_content = "Pregunta,Respuesta\n"
        result = parse_answer_key_csv(csv_content)
        assert result == {}

    def test_lowercase_letters(self):
        csv_content = "Pregunta,Respuesta\n1,A\n2,B\n"
        result = parse_answer_key_csv(csv_content)
        assert result[1] == "a"
        assert result[2] == "b"
```

- [ ] **Paso 3: Correr tests para verificar que fallan**

Comando: `uv run pytest tests/test_answer_key.py -v -k "Parse"`
Esperado: FAIL — `src.answer_key` no existe todavía.

- [ ] **Paso 4: Implementar `parse_answer_key_csv`**

Crear `src/answer_key.py`:

```python
"""Generación de pautas de respuestas para formas randomizadas.

Funciones:
- parse_answer_key_csv: Parsea la pauta original (CSV exportado de Google Sheets)
- generate_answer_key: Genera el mapeo {pregunta → respuesta} para una forma randomizada
- build_answer_key_xlsx: Escribe la pauta como archivo .xlsx
"""

import csv
import io

from openpyxl import Workbook

from src.models import Question


def parse_answer_key_csv(csv_content: str) -> dict[int, str]:
    """Parsea el contenido CSV de la pauta original.

    Formato esperado:
        Pregunta,Respuesta
        1,a
        2,b
        ...

    Args:
        csv_content: String con el contenido CSV completo.

    Returns:
        Dict mapeando número de pregunta a letra de respuesta correcta.
        Ejemplo: {1: "a", 2: "b", 3: "c"}
    """
    reader = csv.DictReader(io.StringIO(csv_content))
    answer_key: dict[int, str] = {}

    for row in reader:
        q_num_raw = row.get("Pregunta", "").strip()
        answer_raw = row.get("Respuesta", "").strip().lower()

        if not q_num_raw or not answer_raw:
            continue

        answer_key[int(q_num_raw)] = answer_raw

    return answer_key
```

- [ ] **Paso 5: Correr tests de parseo**

Comando: `uv run pytest tests/test_answer_key.py -v -k "Parse"`
Esperado: PASS

- [ ] **Paso 6: Escribir tests para generación de pauta de una forma**

```python
# En tests/test_answer_key.py — agregar

from src.models import Alternative, Question
from lxml import etree

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _make_p(text: str):
    p = etree.SubElement(etree.Element("root"), f"{{{NS_W}}}p")
    r = etree.SubElement(p, f"{{{NS_W}}}r")
    t = etree.SubElement(r, f"{{{NS_W}}}t")
    t.text = text
    return p


def _make_question(number: int, original_number: int, alt_mapping: list[tuple[str, str]]) -> Question:
    """Crea una Question con alternativas que tienen letter y original_letter.

    Args:
        number: Número actual (en la forma randomizada).
        original_number: Número original (en el examen original).
        alt_mapping: Lista de (current_letter, original_letter).
    """
    alts = []
    for current, original in alt_mapping:
        alts.append(Alternative(
            letter=current,
            paragraphs=[_make_p(f"{current}) texto")],
            original_letter=original,
        ))
    return Question(
        number=number,
        enunciado_paragraphs=[_make_p(f"{number}- Pregunta")],
        alternatives=alts,
        original_number=original_number,
    )


class TestGenerateAnswerKey:
    """Tests para generate_answer_key."""

    def test_simple_mapping(self):
        """Si Q original 1 tiene respuesta 'a', y 'a' original ahora es 'c',
        la respuesta para la nueva Q debe ser 'c'."""
        original_key = {1: "a", 2: "b"}

        # Q1 nueva = Q1 original, alternativas mezcladas: orig-a→ahora-c
        q1 = _make_question(1, original_number=1, alt_mapping=[
            ("a", "d"), ("b", "b"), ("c", "a"), ("d", "c"),
        ])
        # Q2 nueva = Q2 original, alternativas mezcladas: orig-b→ahora-a
        q2 = _make_question(2, original_number=2, alt_mapping=[
            ("a", "b"), ("b", "c"), ("c", "d"), ("d", "a"),
        ])

        result = generate_answer_key([q1, q2], original_key)
        assert result == {1: "c", 2: "a"}

    def test_preserves_all_questions(self):
        original_key = {1: "a", 2: "b", 3: "c"}
        questions = [
            _make_question(1, 3, [("a", "c"), ("b", "a"), ("c", "d"), ("d", "b")]),
            _make_question(2, 1, [("a", "b"), ("b", "a"), ("c", "d"), ("d", "c")]),
            _make_question(3, 2, [("a", "d"), ("b", "c"), ("c", "b"), ("d", "a")]),
        ]
        result = generate_answer_key(questions, original_key)
        # Q1 nueva (orig Q3, resp "c"): orig "c" → ahora en "a" → respuesta = "a"? No...
        # Q3 original correcta = "c". En Q1 nueva: alternativas orig_letter c→current a? Sí, ("a","c")? Sí, primer tupla.
        # Pero ("a","c") means current="a", original="c". So original "c" maps to current "a".
        assert result[1] == "a"
        # Q2 nueva (orig Q1, resp "a"): orig "a" is in ("b","a") → current "b"
        assert result[2] == "b"
        # Q3 nueva (orig Q2, resp "b"): orig "b" is in ("c","b") → current "c"
        assert result[3] == "c"
        assert len(result) == 3

    def test_missing_original_raises(self):
        """Si una pregunta referencia un original_number que no existe en la pauta, error."""
        original_key = {1: "a"}
        q = _make_question(1, original_number=99, alt_mapping=[
            ("a", "a"), ("b", "b"), ("c", "c"), ("d", "d"),
        ])
        with pytest.raises(KeyError):
            generate_answer_key([q], original_key)
```

- [ ] **Paso 7: Correr tests para verificar que fallan**

Comando: `uv run pytest tests/test_answer_key.py -v -k "Generate"`
Esperado: FAIL — `generate_answer_key` no existe.

- [ ] **Paso 8: Implementar `generate_answer_key`**

Agregar a `src/answer_key.py`:

```python
def generate_answer_key(
    questions: list[Question],
    original_key: dict[int, str],
) -> dict[int, str]:
    """Genera la pauta de respuestas para una forma randomizada.

    Para cada pregunta:
    1. Buscar la respuesta correcta original usando original_number
    2. Encontrar qué alternativa tiene esa original_letter
    3. La letra actual de esa alternativa es la nueva respuesta correcta

    Args:
        questions: Lista de preguntas randomizadas (con original_number/original_letter).
        original_key: Pauta original {número_pregunta: letra_correcta}.

    Returns:
        Dict {nuevo_número: nueva_letra_correcta}.

    Raises:
        KeyError: Si original_number no existe en original_key.
        ValueError: Si la letra correcta original no se encuentra en las alternativas.
    """
    new_key: dict[int, str] = {}

    for q in questions:
        correct_original_letter = original_key[q.original_number]

        new_letter = None
        for alt in q.alternatives:
            if alt.original_letter == correct_original_letter:
                new_letter = alt.letter
                break

        if new_letter is None:
            raise ValueError(
                f"Q{q.number} (original Q{q.original_number}): "
                f"no se encontró alternativa con original_letter='{correct_original_letter}'. "
                f"Alternativas: {[(a.letter, a.original_letter) for a in q.alternatives]}"
            )

        new_key[q.number] = new_letter

    return new_key
```

- [ ] **Paso 9: Correr tests de generación**

Comando: `uv run pytest tests/test_answer_key.py -v -k "Generate"`
Esperado: PASS

- [ ] **Paso 10: Escribir tests para escritura de .xlsx**

```python
# En tests/test_answer_key.py — agregar

import openpyxl


class TestBuildAnswerKeyXlsx:
    """Tests para escritura del archivo .xlsx de pauta."""

    def test_creates_valid_xlsx(self):
        answer_key = {1: "a", 2: "b", 3: "c"}
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            build_answer_key_xlsx(answer_key, path)

            assert os.path.exists(path)
            assert os.path.getsize(path) > 0

            wb = openpyxl.load_workbook(path)
            ws = wb.active
            assert ws.cell(1, 1).value == "Pregunta"
            assert ws.cell(1, 2).value == "Respuesta"
            assert ws.cell(2, 1).value == 1
            assert ws.cell(2, 2).value == "a"
            assert ws.cell(3, 1).value == 2
            assert ws.cell(3, 2).value == "b"
            assert ws.cell(4, 1).value == 3
            assert ws.cell(4, 2).value == "c"
        finally:
            os.unlink(path)

    def test_sorted_by_question_number(self):
        answer_key = {3: "c", 1: "a", 2: "b"}
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            build_answer_key_xlsx(answer_key, path)
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            assert ws.cell(2, 1).value == 1
            assert ws.cell(3, 1).value == 2
            assert ws.cell(4, 1).value == 3
        finally:
            os.unlink(path)

    def test_36_questions(self):
        answer_key = {i: chr(ord("a") + (i - 1) % 4) for i in range(1, 37)}
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            build_answer_key_xlsx(answer_key, path)
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            assert ws.max_row == 37  # 1 header + 36 data
            assert ws.cell(37, 1).value == 36
            assert ws.cell(37, 2).value == "d"
        finally:
            os.unlink(path)
```

- [ ] **Paso 11: Correr tests para verificar que fallan**

Comando: `uv run pytest tests/test_answer_key.py -v -k "Build"`
Esperado: FAIL — `build_answer_key_xlsx` no existe.

- [ ] **Paso 12: Implementar `build_answer_key_xlsx`**

Agregar a `src/answer_key.py`:

```python
def build_answer_key_xlsx(
    answer_key: dict[int, str],
    output_path: str,
) -> None:
    """Escribe la pauta de respuestas como archivo .xlsx.

    Formato:
        A1: "Pregunta", B1: "Respuesta"
        A2: 1, B2: "a"
        A3: 2, B3: "b"
        ...

    Las filas se ordenan por número de pregunta ascendente.

    Args:
        answer_key: Dict {número_pregunta: letra_correcta}.
        output_path: Ruta donde guardar el .xlsx.
    """
    wb = Workbook()
    ws = wb.active

    ws.cell(row=1, column=1, value="Pregunta")
    ws.cell(row=1, column=2, value="Respuesta")

    for row_idx, q_num in enumerate(sorted(answer_key.keys()), start=2):
        ws.cell(row=row_idx, column=1, value=q_num)
        ws.cell(row=row_idx, column=2, value=answer_key[q_num])

    wb.save(output_path)
```

- [ ] **Paso 13: Correr todos los tests de answer_key**

Comando: `uv run pytest tests/test_answer_key.py -v`
Esperado: PASS

- [ ] **Paso 14: Correr suite completa**

Comando: `uv run pytest tests/ -v`
Esperado: PASS

- [ ] **Paso 15: Commit**

```bash
git add src/answer_key.py tests/test_answer_key.py pyproject.toml
git commit -m "feat: add answer key parsing, generation, and xlsx writing"
```

---

### Tarea 4: Agregar soporte en drive_client para Google Sheets y .xlsx

**Archivos:**
- Modificar: `src/drive_client.py`

El archivo de claves original está almacenado como un **Google Sheet nativo** (`application/vnd.google-apps.spreadsheet`), no como un `.xlsx`. La descarga requiere `export_media` (exportar como CSV) en vez de `get_media` (descarga directa). Además, la subida de pautas generadas son archivos `.xlsx` con un MIME type distinto a `.docx`.

- [ ] **Paso 1: Implementar `download_sheet_as_csv`**

Agregar a `src/drive_client.py`:

```python
def download_sheet_as_csv(drive_path: str) -> str:
    """Descarga un Google Sheet nativo y lo retorna como string CSV.

    Los Google Sheets nativos no se pueden descargar con get_media —
    requieren export_media con un MIME type de exportación.

    Args:
        drive_path: Ruta en Drive (ej: "Guías/Claves Ondas 2026. Ev. 1 F1").
                    Sin extensión — los Google Sheets nativos no tienen.

    Returns:
        String con el contenido CSV del sheet.

    Raises:
        FileNotFoundError: Si no se encuentra el archivo en Drive.
    """
    service = _get_service()
    file_id = _resolve_path(service, drive_path)

    request = service.files().export_media(
        fileId=file_id,
        mimeType="text/csv",
    )

    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    csv_content = buffer.getvalue().decode("utf-8")
    print(f"✅ Pauta descargada: {drive_path}")
    return csv_content
```

Nota: agregar `import io` al inicio del archivo si no está.

- [ ] **Paso 2: Implementar `upload_xlsx`**

Agregar a `src/drive_client.py`:

```python
def upload_xlsx(local_path: str, drive_folder_path: str, filename: str) -> str:
    """Sube un archivo .xlsx a una carpeta en Google Drive.

    Args:
        local_path: Ruta local al .xlsx.
        drive_folder_path: Ruta de la carpeta destino en Drive.
        filename: Nombre del archivo en Drive.

    Returns:
        ID del archivo creado en Drive.
    """
    service = _get_service()

    if drive_folder_path:
        parent_id = _resolve_path(service, drive_folder_path)
    else:
        parent_id = "root"

    file_metadata = {"name": filename, "parents": [parent_id]}
    media = MediaFileUpload(
        local_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    file = (
        service.files()
        .create(body=file_metadata, media_body=media, fields="id")
        .execute()
    )

    print(f"✅ Pauta subida: {filename}")
    return file.get("id")
```

- [ ] **Paso 3: Correr suite completa**

Comando: `uv run pytest tests/ -v`
Esperado: PASS — los cambios en drive_client no rompen tests existentes.

- [ ] **Paso 4: Commit**

```bash
git add src/drive_client.py
git commit -m "feat: add download_sheet_as_csv and upload_xlsx to drive_client"
```

---

### Tarea 5: Integrar generación de pautas en main.py

**Archivos:**
- Modificar: `src/main.py`
- Modificar: `tests/test_main.py`

El flujo principal debe:
1. Derivar la ruta del archivo de claves desde el nombre de la prueba
2. Descargar la pauta original (Google Sheet → CSV)
3. Después de generar cada forma, generar su pauta y subirla

- [ ] **Paso 1: Escribir test para derivación del nombre de claves**

```python
# En tests/test_main.py — agregar

from src.main import derive_answer_key_path


class TestDeriveAnswerKeyPath:
    def test_simple(self):
        path = derive_answer_key_path("Guías/Ondas 2026. Ev. 1 F1.docx")
        assert path == "Guías/Claves Ondas 2026. Ev. 1 F1"

    def test_no_folder(self):
        path = derive_answer_key_path("Prueba F1.docx")
        assert path == "Claves Prueba F1"

    def test_deep_folder(self):
        path = derive_answer_key_path("A/B/C/Test F3.docx")
        assert path == "A/B/C/Claves Test F3"
```

- [ ] **Paso 2: Correr tests para verificar que fallan**

Comando: `uv run pytest tests/test_main.py -v -k "Derive"`
Esperado: FAIL — `derive_answer_key_path` no existe.

- [ ] **Paso 3: Implementar `derive_answer_key_path`**

Agregar a `src/main.py`:

```python
def derive_answer_key_path(drive_path: str) -> str:
    """Deriva la ruta del archivo de claves a partir de la ruta de la prueba.

    Convención: en la misma carpeta, con prefijo "Claves " y sin extensión
    (porque es un Google Sheet nativo).

    Ejemplo:
        "Guías/Ondas 2026. Ev. 1 F1.docx" → "Guías/Claves Ondas 2026. Ev. 1 F1"

    Args:
        drive_path: Ruta completa al .docx de la prueba en Drive.

    Returns:
        Ruta al Google Sheet de claves en Drive (sin extensión).
    """
    path_parts = drive_path.strip("/").split("/")
    filename = path_parts[-1]
    folder = "/".join(path_parts[:-1])

    # Remover extensión .docx
    base_name = filename.rsplit(".docx", 1)[0]

    # Agregar prefijo "Claves "
    claves_name = f"Claves {base_name}"

    if folder:
        return f"{folder}/{claves_name}"
    return claves_name
```

- [ ] **Paso 4: Correr tests**

Comando: `uv run pytest tests/test_main.py -v -k "Derive"`
Esperado: PASS

- [ ] **Paso 5: Integrar en `main()`**

Modificar `src/main.py` — actualizar imports y la función `main()`:

```python
# Agregar imports
from src.answer_key import parse_answer_key_csv, generate_answer_key, build_answer_key_xlsx
from src.drive_client import download_docx, upload_docx, download_sheet_as_csv, upload_xlsx
```

En la función `main()`, agregar después del parseo del documento y antes de la generación de formas:

```python
        # 2.5. Descargar pauta original
        answer_key_path = derive_answer_key_path(drive_path)
        print(f"📋 Descargando pauta original: {answer_key_path}")
        try:
            csv_content = download_sheet_as_csv(answer_key_path)
            original_key = parse_answer_key_csv(csv_content)
            print(f"   Encontradas {len(original_key)} respuestas.\n")
        except FileNotFoundError:
            print(
                f"⚠️  No se encontró pauta en '{answer_key_path}'. "
                f"Se generarán las formas sin pautas.\n"
            )
            original_key = None
```

Dentro del loop de generación de formas, después de `build_docx` y `upload_docx`, agregar:

```python
            # Generar y subir pauta de esta forma
            if original_key is not None:
                new_key = generate_answer_key(form_questions, original_key)

                key_filename = f"Claves {prefix}{new_form_number}.xlsx"
                key_path = os.path.join(temp_dir, key_filename)
                build_answer_key_xlsx(new_key, key_path)

                print(f"📋 Subiendo pauta: {key_filename}")
                upload_xlsx(key_path, drive_folder, key_filename)
                os.unlink(key_path)
```

- [ ] **Paso 6: Correr suite completa**

Comando: `uv run pytest tests/ -v`
Esperado: PASS

- [ ] **Paso 7: Commit**

```bash
git add src/main.py tests/test_main.py
git commit -m "feat: integrate answer key generation into main pipeline"
```

---

### Tarea 6: Test de integración end-to-end

**Archivos:**
- Modificar: `tests/test_answer_key.py`

Un test que simula el flujo completo: parsear examen → generar forma → generar pauta → verificar que la pauta es consistente.

- [ ] **Paso 1: Escribir test de integración**

```python
# En tests/test_answer_key.py — agregar

from src.parser import parse_docx
from src.randomizer import generate_forms


class TestAnswerKeyIntegration:
    """Test end-to-end: examen + pauta original → forma randomizada → pauta nueva."""

    def _create_exam_with_known_answers(self, path: str):
        """Crea un .docx donde la respuesta correcta siempre es 'a'."""
        from docx import Document

        doc = Document()
        doc.add_paragraph("Prueba de Física")
        doc.add_paragraph("Forma 1")

        qt = doc.add_table(rows=3, cols=2)
        for i in range(6):
            row_idx = i // 2
            col_idx = i % 2
            cell = qt.cell(row_idx, col_idx)
            cell.paragraphs[0].text = f"{i + 1}- Pregunta {i + 1}"
            cell.add_paragraph(f"a) Correcta {i + 1}")
            cell.add_paragraph(f"b) Incorrecta B")
            cell.add_paragraph(f"c) Incorrecta C")
            cell.add_paragraph(f"d) Incorrecta D")

        doc.save(path)

    def test_answer_key_tracks_correct_answer_through_shuffle(self):
        """La respuesta correcta debe seguir apuntando al texto correcto."""
        with tempfile.TemporaryDirectory() as tmpdir:
            source = os.path.join(tmpdir, "original.docx")
            self._create_exam_with_known_answers(source)

            exam = parse_docx(source)
            original_key = {i: "a" for i in range(1, 7)}  # Todas las respuestas son "a"

            import random
            random.seed(123)
            forms = generate_forms(exam, n=1)
            form = forms[0]

            new_key = generate_answer_key(form, original_key)

            # Verificar: para cada pregunta en la forma, la alternativa indicada
            # por la pauta nueva debe contener "Correcta" en su texto
            for q in form:
                correct_letter = new_key[q.number]
                correct_alt = next(a for a in q.alternatives if a.letter == correct_letter)
                assert "Correcta" in correct_alt.text, (
                    f"Q{q.number} (orig Q{q.original_number}): "
                    f"pauta dice '{correct_letter}' pero texto es '{correct_alt.text}'"
                )

    def test_answer_key_xlsx_roundtrip(self):
        """Generar xlsx → leerlo → debe coincidir con el dict."""
        answer_key = {1: "c", 2: "a", 3: "d", 4: "b", 5: "a", 6: "c"}
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            build_answer_key_xlsx(answer_key, path)

            # Leer de vuelta
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            read_back = {}
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
                q_num, letter = row
                read_back[q_num] = letter

            assert read_back == answer_key
        finally:
            os.unlink(path)
```

- [ ] **Paso 2: Correr tests**

Comando: `uv run pytest tests/test_answer_key.py -v -k "Integration"`
Esperado: PASS

- [ ] **Paso 3: Correr suite completa**

Comando: `uv run pytest tests/ -v`
Esperado: PASS

- [ ] **Paso 4: Commit**

```bash
git add tests/test_answer_key.py
git commit -m "test: add end-to-end integration test for answer key generation"
```

---

## Resumen de tareas

| #   | Tarea                                             | Archivos                                        | Dependencias |
| --- | ------------------------------------------------- | ----------------------------------------------- | ------------ |
| 1   | Agregar campos de rastreo a modelos               | `src/models.py`, `tests/test_models.py`         | —            |
| 2   | Preservar originales en randomizer                | `tests/test_randomizer.py`                      | 1            |
| 3   | Crear módulo answer_key (parse + generate + xlsx) | `src/answer_key.py`, `tests/test_answer_key.py`, `pyproject.toml` | 1        |
| 4   | Soporte Drive para Sheets y xlsx                  | `src/drive_client.py`                           | —            |
| 5   | Integrar en main.py                               | `src/main.py`, `tests/test_main.py`             | 3, 4         |
| 6   | Test de integración end-to-end                    | `tests/test_answer_key.py`                      | 1, 2, 3      |

## Decisiones de diseño clave

1. **`original_letter` / `original_number` con defaults retrocompatibles:** Usar `__post_init__` para copiar desde `letter`/`number` cuando no se especifica. Esto asegura que todo el código existente sigue funcionando sin cambios — el parser no necesita setear estos campos explícitamente.

2. **CSV en vez de xlsx para la pauta original:** El archivo de claves es un Google Sheet nativo (no un .xlsx). La API de Drive requiere `export_media` para estos archivos. Exportar como CSV es más simple y no requiere dependencias extra para lectura (solo `csv` de stdlib). `openpyxl` se usa solo para **escribir** las pautas generadas.

3. **Pauta opcional (graceful degradation):** Si no se encuentra el archivo de claves en Drive, el programa imprime un warning y genera las formas sin pautas. Esto mantiene el flujo existente intacto para casos donde no hay pauta.

4. **Nombre de pauta generada:** `Claves <prefix><form_number>.xlsx` — coherente con la convención del original pero como archivo `.xlsx` real (no Google Sheet nativo). Esto permite que el usuario lo abra directamente con Excel/Sheets.

## Risks

- **Google Sheets con múltiples hojas:** El export a CSV solo exporta la primera hoja. Si el sheet de claves tiene varias hojas, solo se lee la primera. Mitigación: documentarlo y verificar que el sheet real solo tiene una hoja.
- **Variaciones en headers del CSV:** Si el Google Sheet tiene headers diferentes ("Pregunta"/"Respuesta" vs otra cosa), el parseo falla silenciosamente. Mitigación: validar que las columnas esperadas existen y dar error claro.