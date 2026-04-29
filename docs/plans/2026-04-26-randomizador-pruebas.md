# Randomizador de Pruebas — Plan de Implementación

> **Para agentes:** Usá `@rules/skills/subagents.md` (recomendado) o `@rules/skills/executing-plans.md` para implementar este plan tarea por tarea. Los pasos usan sintaxis checkbox (`- [ ]`) para tracking.

**Goal:** Crear un script en Python que lea un archivo .docx de pruebas de alternativas desde Google Drive, randomice el orden de preguntas y alternativas, y genere N formas distintas como nuevos archivos .docx en la misma carpeta.

**Arquitectura:** El sistema se compone de 4 módulos: (1) conexión y descarga/subida a Google Drive vía API, (2) parser del .docx que extrae preguntas con sus enunciados, imágenes opcionales y alternativas, (3) randomizador que genera N permutaciones distintas de preguntas y alternativas, y (4) generador de .docx que reconstruye el documento con el nuevo orden. Un script principal `main.py` orquesta todo.

**Tech Stack:** Python 3.10+, `python-docx` (manipulación de .docx), `google-api-python-client` + `google-auth-oauthlib` (Google Drive API), `random` (stdlib para shuffle).

---

## Estructura de archivos

| Archivo | Responsabilidad |
|---|---|
| `requirements.txt` | Dependencias del proyecto |
| `src/main.py` | Script principal — orquesta el flujo completo, recibe argumentos del usuario |
| `src/drive_client.py` | Conexión a Google Drive: autenticación OAuth2, descarga de .docx, subida de archivos generados |
| `src/parser.py` | Parsea el .docx: extrae preguntas (número + enunciado + imágenes) y alternativas (letra + texto) |
| `src/randomizer.py` | Genera N permutaciones distintas de preguntas y alternativas |
| `src/doc_builder.py` | Construye nuevos archivos .docx a partir de las preguntas reordenadas |
| `src/models.py` | Dataclasses: `Question`, `Alternative`, `Exam` |
| `tests/test_parser.py` | Tests del parser |
| `tests/test_randomizer.py` | Tests del randomizador |
| `tests/test_doc_builder.py` | Tests del generador de documentos |
| `credentials/README.md` | Instrucciones para obtener credenciales de Google Drive API |

---

### Tarea 1: Setup del proyecto y dependencias

**Archivos:**
- Crear: `requirements.txt`
- Crear: `src/__init__.py`
- Crear: `tests/__init__.py`
- Crear: `credentials/README.md`
- Crear: `.gitignore`

- [ ] **Paso 1: Crear `requirements.txt`**

```
python-docx==1.1.2
google-api-python-client==2.131.0
google-auth-oauthlib==1.2.0
google-auth-httplib2==0.2.0
```

- [ ] **Paso 2: Crear archivos `__init__.py` vacíos**

Crear `src/__init__.py` y `tests/__init__.py` como archivos vacíos.

- [ ] **Paso 3: Crear `.gitignore`**

```
__pycache__/
*.pyc
.env
credentials/client_secret.json
credentials/token.json
venv/
*.egg-info/
dist/
build/
```

- [ ] **Paso 4: Crear `credentials/README.md`**

```markdown
# Credenciales de Google Drive API

## Pasos para obtener credenciales:

1. Ir a [Google Cloud Console](https://console.cloud.google.com/)
2. Crear un proyecto nuevo (o usar uno existente)
3. Habilitar la **Google Drive API** en "APIs & Services" > "Library"
4. Ir a "APIs & Services" > "Credentials"
5. Crear credenciales tipo **OAuth 2.0 Client ID** (tipo "Desktop App")
6. Descargar el JSON y guardarlo como `credentials/client_secret.json`

## Archivos generados

- `client_secret.json` — credenciales OAuth2 (NO subir a git)
- `token.json` — se genera automáticamente en la primera ejecución (NO subir a git)
```

- [ ] **Paso 5: Instalar dependencias**

Comando: `python -m venv venv && source venv/bin/activate && pip install -r requirements.txt`

- [ ] **Paso 6: Commit**

```bash
git init
git add requirements.txt src/__init__.py tests/__init__.py credentials/README.md .gitignore
git commit -m "chore: setup inicial del proyecto con dependencias"
```

---

### Tarea 2: Modelos de datos

**Archivos:**
- Crear: `src/models.py`
- Crear: `tests/test_models.py`

- [ ] **Paso 1: Escribir el test para los modelos**

```python
# tests/test_models.py
from src.models import Alternative, Question, Exam


def test_alternative_creation():
    alt = Alternative(letter="a", text="Velocidad de la luz", image=None)
    assert alt.letter == "a"
    assert alt.text == "Velocidad de la luz"
    assert alt.image is None


def test_question_creation():
    alts = [
        Alternative(letter="a", text="Opción A", image=None),
        Alternative(letter="b", text="Opción B", image=None),
    ]
    q = Question(number=1, text="¿Cuál es correcta?", alternatives=alts, images=[])
    assert q.number == 1
    assert len(q.alternatives) == 2


def test_exam_creation():
    alts = [
        Alternative(letter="a", text="Opción A", image=None),
    ]
    q = Question(number=1, text="Pregunta 1", alternatives=alts, images=[])
    exam = Exam(title="Prueba de Ondas", questions=[q])
    assert exam.title == "Prueba de Ondas"
    assert len(exam.questions) == 1
```

- [ ] **Paso 2: Correr el test para verificar que falla**

Comando: `python -m pytest tests/test_models.py -v`
Esperado: FAIL con "ModuleNotFoundError" o "ImportError"

- [ ] **Paso 3: Implementar los modelos**

```python
# src/models.py
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Alternative:
    """Una alternativa de una pregunta de selección múltiple.

    Attributes:
        letter: Letra original de la alternativa (a, b, c, d, e).
        text: Texto de la alternativa.
        image: Bytes de la imagen asociada (si existe). Tupla (bytes, content_type).
    """
    letter: str
    text: str
    image: Optional[tuple[bytes, str]] = None


@dataclass
class Question:
    """Una pregunta completa con enunciado y alternativas.

    Attributes:
        number: Número original de la pregunta (1, 2, 3...).
        text: Texto del enunciado de la pregunta.
        alternatives: Lista de alternativas de la pregunta.
        images: Lista de imágenes del enunciado. Cada una es (bytes, content_type).
    """
    number: int
    text: str
    alternatives: list[Alternative] = field(default_factory=list)
    images: list[tuple[bytes, str]] = field(default_factory=list)


@dataclass
class Exam:
    """Un examen completo con título y preguntas.

    Attributes:
        title: Título del examen (nombre del archivo sin extensión).
        questions: Lista de preguntas del examen.
        header_paragraphs: Párrafos que aparecen antes de la primera pregunta (instrucciones, etc).
    """
    title: str
    questions: list[Question] = field(default_factory=list)
    header_paragraphs: list[str] = field(default_factory=list)
```

- [ ] **Paso 4: Correr el test para verificar que pasa**

Comando: `python -m pytest tests/test_models.py -v`
Esperado: PASS (3 tests)

- [ ] **Paso 5: Commit**

```bash
git add src/models.py tests/test_models.py
git commit -m "feat: agregar modelos de datos Question, Alternative, Exam"
```

---

### Tarea 3: Parser del documento .docx

**Archivos:**
- Crear: `src/parser.py`
- Crear: `tests/test_parser.py`
- Crear: `tests/fixtures/sample_exam.docx` (generado en el test setup)

- [ ] **Paso 1: Escribir el test del parser**

```python
# tests/test_parser.py
import os
import tempfile
from docx import Document
from docx.shared import Inches
from src.parser import parse_docx
from src.models import Exam


def create_sample_docx(path: str):
    """Crea un .docx de prueba con estructura de examen."""
    doc = Document()
    doc.add_paragraph("Prueba de Física - Ondas")
    doc.add_paragraph("Nombre: ________________")
    doc.add_paragraph("")
    doc.add_paragraph("1- ¿Cuál es la velocidad de la luz en el vacío?")
    doc.add_paragraph("a) 3 x 10^8 m/s")
    doc.add_paragraph("b) 3 x 10^6 m/s")
    doc.add_paragraph("c) 3 x 10^10 m/s")
    doc.add_paragraph("d) 3 x 10^4 m/s")
    doc.add_paragraph("")
    doc.add_paragraph("2- ¿Qué tipo de onda es el sonido?")
    doc.add_paragraph("a) Transversal")
    doc.add_paragraph("b) Longitudinal")
    doc.add_paragraph("c) Electromagnética")
    doc.add_paragraph("d) Estacionaria")
    doc.add_paragraph("")
    doc.add_paragraph("3- ¿Cuál es la unidad de frecuencia?")
    doc.add_paragraph("a) Metro")
    doc.add_paragraph("b) Segundo")
    doc.add_paragraph("c) Hertz")
    doc.add_paragraph("d) Newton")
    doc.save(path)


def test_parse_docx_extracts_questions():
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name
    try:
        create_sample_docx(path)
        exam = parse_docx(path)

        assert isinstance(exam, Exam)
        assert len(exam.questions) == 3

        # Verificar primera pregunta
        q1 = exam.questions[0]
        assert q1.number == 1
        assert "velocidad de la luz" in q1.text.lower()
        assert len(q1.alternatives) == 4
        assert q1.alternatives[0].letter == "a"
        assert "3 x 10^8" in q1.alternatives[0].text

        # Verificar segunda pregunta
        q2 = exam.questions[1]
        assert q2.number == 2
        assert "sonido" in q2.text.lower()
        assert len(q2.alternatives) == 4

        # Verificar tercera pregunta
        q3 = exam.questions[2]
        assert q3.number == 3
        assert len(q3.alternatives) == 4
    finally:
        os.unlink(path)


def test_parse_docx_extracts_header():
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name
    try:
        create_sample_docx(path)
        exam = parse_docx(path)

        # Los párrafos antes de la primera pregunta deben estar en header_paragraphs
        assert len(exam.header_paragraphs) >= 1
        assert any("Prueba de Física" in h for h in exam.header_paragraphs)
    finally:
        os.unlink(path)


def test_parse_docx_handles_multiline_enunciado():
    """Test que el parser maneja enunciados que ocupan más de una línea."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name
    try:
        doc = Document()
        doc.add_paragraph("1- Considere el siguiente fenómeno físico:")
        doc.add_paragraph("Una onda viaja por una cuerda tensa.")
        doc.add_paragraph("¿Qué tipo de onda es?")
        doc.add_paragraph("a) Transversal")
        doc.add_paragraph("b) Longitudinal")
        doc.add_paragraph("c) Electromagnética")
        doc.add_paragraph("d) Mecánica")
        doc.save(path)

        exam = parse_docx(path)
        assert len(exam.questions) == 1
        q = exam.questions[0]
        assert q.number == 1
        assert "cuerda tensa" in q.text.lower()
        assert len(q.alternatives) == 4
    finally:
        os.unlink(path)
```

- [ ] **Paso 2: Correr el test para verificar que falla**

Comando: `python -m pytest tests/test_parser.py -v`
Esperado: FAIL con "ImportError: cannot import name 'parse_docx'"

- [ ] **Paso 3: Implementar el parser**

```python
# src/parser.py
"""Parser de archivos .docx con formato de pruebas de alternativas.

Formato esperado:
- Preguntas empiezan con "N-" (ej: "1-", "2-", "10-")
- Alternativas empiezan con "x)" (ej: "a)", "b)", "c)")
- Las preguntas pueden tener imágenes embebidas en el enunciado
- Todo lo que aparece antes de la primera pregunta se considera header
"""
import re
import io
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from src.models import Alternative, Question, Exam

# Regex para detectar inicio de pregunta: "N-" donde N es uno o más dígitos
QUESTION_PATTERN = re.compile(r"^(\d+)\s*[-–—]\s*(.*)$")

# Regex para detectar inicio de alternativa: "x)" donde x es una letra minúscula
ALTERNATIVE_PATTERN = re.compile(r"^([a-zA-Z])\)\s*(.*)$")


def _extract_images_from_paragraph(paragraph) -> list[tuple[bytes, str]]:
    """Extrae imágenes embebidas en un párrafo del documento.

    Args:
        paragraph: Un objeto Paragraph de python-docx.

    Returns:
        Lista de tuplas (image_bytes, content_type).
    """
    images = []
    for run in paragraph.runs:
        # Buscar imágenes inline (blipFill)
        drawing_elements = run._element.findall(
            ".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
        )
        for blip in drawing_elements:
            embed_attr = blip.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
            )
            if embed_attr:
                try:
                    rel = paragraph.part.rels[embed_attr]
                    image_blob = rel.target_part.blob
                    content_type = rel.target_part.content_type
                    images.append((image_blob, content_type))
                except (KeyError, AttributeError):
                    pass
    return images


def parse_docx(file_path: str) -> Exam:
    """Parsea un archivo .docx de prueba de alternativas.

    Args:
        file_path: Ruta al archivo .docx.

    Returns:
        Un objeto Exam con las preguntas, alternativas e imágenes extraídas.
    """
    doc = Document(file_path)
    questions: list[Question] = []
    header_paragraphs: list[str] = []
    current_question: Question | None = None
    found_first_question = False

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        images = _extract_images_from_paragraph(paragraph)

        # Intentar match con patrón de pregunta
        q_match = QUESTION_PATTERN.match(text)
        if q_match:
            # Si había una pregunta en curso, guardarla
            if current_question is not None:
                questions.append(current_question)

            found_first_question = True
            number = int(q_match.group(1))
            enunciado = q_match.group(2).strip()
            current_question = Question(
                number=number,
                text=enunciado,
                alternatives=[],
                images=list(images),
            )
            continue

        # Intentar match con patrón de alternativa
        a_match = ALTERNATIVE_PATTERN.match(text)
        if a_match and current_question is not None:
            letter = a_match.group(1).lower()
            alt_text = a_match.group(2).strip()
            alt_image = images[0] if images else None
            current_question.alternatives.append(
                Alternative(letter=letter, text=alt_text, image=alt_image)
            )
            continue

        # Si no es ni pregunta ni alternativa
        if not found_first_question:
            # Es parte del header (instrucciones, título, nombre, etc.)
            if text:  # Ignorar líneas vacías del header
                header_paragraphs.append(text)
        elif current_question is not None and text:
            # Es continuación del enunciado de la pregunta actual
            # (enunciados multilinea)
            current_question.text += "\n" + text
            current_question.images.extend(images)

    # No olvidar la última pregunta
    if current_question is not None:
        questions.append(current_question)

    title = header_paragraphs[0] if header_paragraphs else "Examen"

    return Exam(
        title=title,
        questions=questions,
        header_paragraphs=header_paragraphs,
    )
```

- [ ] **Paso 4: Correr el test para verificar que pasa**

Comando: `python -m pytest tests/test_parser.py -v`
Esperado: PASS (3 tests)

- [ ] **Paso 5: Commit**

```bash
git add src/parser.py tests/test_parser.py
git commit -m "feat: implementar parser de .docx con preguntas y alternativas"
```

---

### Tarea 4: Randomizador de preguntas y alternativas

**Archivos:**
- Crear: `src/randomizer.py`
- Crear: `tests/test_randomizer.py`

- [ ] **Paso 1: Escribir el test del randomizador**

```python
# tests/test_randomizer.py
import copy
from src.models import Alternative, Question, Exam
from src.randomizer import generate_forms


def _make_exam() -> Exam:
    """Crea un examen de prueba con 5 preguntas y 4 alternativas cada una."""
    questions = []
    for i in range(1, 6):
        alts = [
            Alternative(letter="a", text=f"Alt A de P{i}", image=None),
            Alternative(letter="b", text=f"Alt B de P{i}", image=None),
            Alternative(letter="c", text=f"Alt C de P{i}", image=None),
            Alternative(letter="d", text=f"Alt D de P{i}", image=None),
        ]
        questions.append(
            Question(number=i, text=f"Pregunta {i}", alternatives=alts, images=[])
        )
    return Exam(title="Examen Test", questions=questions, header_paragraphs=["Header"])


def test_generate_forms_returns_correct_count():
    exam = _make_exam()
    forms = generate_forms(exam, n=3)
    assert len(forms) == 3


def test_generate_forms_preserves_all_questions():
    exam = _make_exam()
    forms = generate_forms(exam, n=2)
    for form in forms:
        assert len(form.questions) == 5
        # Verificar que todas las preguntas originales están presentes
        original_texts = {q.text for q in exam.questions}
        form_texts = {q.text for q in form.questions}
        assert original_texts == form_texts


def test_generate_forms_renumbers_questions():
    exam = _make_exam()
    forms = generate_forms(exam, n=2)
    for form in forms:
        # Las preguntas deben estar numeradas 1, 2, 3, 4, 5
        numbers = [q.number for q in form.questions]
        assert numbers == [1, 2, 3, 4, 5]


def test_generate_forms_relabels_alternatives():
    exam = _make_exam()
    forms = generate_forms(exam, n=2)
    for form in forms:
        for q in form.questions:
            letters = [a.letter for a in q.alternatives]
            assert letters == ["a", "b", "c", "d"]


def test_generate_forms_shuffles_order():
    """Al menos una forma debería tener un orden distinto al original."""
    exam = _make_exam()
    # Con 10 formas es prácticamente imposible que todas tengan el mismo orden
    forms = generate_forms(exam, n=10)
    original_order = [q.text for q in exam.questions]
    any_different = any(
        [q.text for q in form.questions] != original_order for form in forms
    )
    assert any_different, "Ninguna forma tiene un orden diferente al original"


def test_generate_forms_shuffles_alternatives():
    """Al menos una pregunta en alguna forma debería tener alternativas reordenadas."""
    exam = _make_exam()
    forms = generate_forms(exam, n=10)
    original_alt_texts = [a.text for a in exam.questions[0].alternatives]
    any_different_alts = False
    for form in forms:
        for q in form.questions:
            alt_texts = [a.text for a in q.alternatives]
            if alt_texts != original_alt_texts:
                any_different_alts = True
                break
    assert any_different_alts, "Ninguna alternativa fue reordenada"


def test_generate_forms_preserves_header():
    exam = _make_exam()
    forms = generate_forms(exam, n=2)
    for form in forms:
        assert form.header_paragraphs == exam.header_paragraphs


def test_generate_forms_does_not_mutate_original():
    exam = _make_exam()
    original_questions = [q.text for q in exam.questions]
    original_alts = [a.text for q in exam.questions for a in q.alternatives]
    generate_forms(exam, n=5)
    assert [q.text for q in exam.questions] == original_questions
    assert [a.text for q in exam.questions for a in q.alternatives] == original_alts
```

- [ ] **Paso 2: Correr el test para verificar que falla**

Comando: `python -m pytest tests/test_randomizer.py -v`
Esperado: FAIL con "ImportError: cannot import name 'generate_forms'"

- [ ] **Paso 3: Implementar el randomizador**

```python
# src/randomizer.py
"""Generador de formas aleatorias a partir de un examen.

Produce N versiones del examen con preguntas y alternativas en orden aleatorio.
Cada forma mantiene todas las preguntas y alternativas originales, pero reordena
tanto el orden de las preguntas como el de las alternativas dentro de cada pregunta.
"""
import copy
import random
from src.models import Alternative, Question, Exam


def _shuffle_alternatives(alternatives: list[Alternative]) -> list[Alternative]:
    """Reordena las alternativas y reasigna las letras a, b, c, d...

    Args:
        alternatives: Lista de alternativas originales.

    Returns:
        Nueva lista de alternativas con orden aleatorio y letras reasignadas.
    """
    shuffled = copy.deepcopy(alternatives)
    random.shuffle(shuffled)
    for i, alt in enumerate(shuffled):
        alt.letter = chr(ord("a") + i)
    return shuffled


def _shuffle_questions(questions: list[Question]) -> list[Question]:
    """Reordena las preguntas, reasigna números, y shufflea sus alternativas.

    Args:
        questions: Lista de preguntas originales.

    Returns:
        Nueva lista de preguntas con orden aleatorio, números reasignados,
        y alternativas reordenadas.
    """
    shuffled = copy.deepcopy(questions)
    random.shuffle(shuffled)
    for i, q in enumerate(shuffled):
        q.number = i + 1
        q.alternatives = _shuffle_alternatives(q.alternatives)
    return shuffled


def generate_forms(exam: Exam, n: int) -> list[Exam]:
    """Genera N formas aleatorias del examen.

    Cada forma tiene las mismas preguntas y alternativas que el original,
    pero en orden distinto. Las preguntas se renumeran (1, 2, 3...) y
    las alternativas se re-etiquetan (a, b, c...).

    Args:
        exam: El examen original parseado.
        n: Número de formas a generar.

    Returns:
        Lista de N exámenes con preguntas y alternativas reordenadas.

    Raises:
        ValueError: Si n < 1.
    """
    if n < 1:
        raise ValueError("El número de formas debe ser al menos 1")

    forms: list[Exam] = []
    for i in range(n):
        shuffled_questions = _shuffle_questions(exam.questions)
        form = Exam(
            title=exam.title,
            questions=shuffled_questions,
            header_paragraphs=list(exam.header_paragraphs),
        )
        forms.append(form)

    return forms
```

- [ ] **Paso 4: Correr el test para verificar que pasa**

Comando: `python -m pytest tests/test_randomizer.py -v`
Esperado: PASS (8 tests)

- [ ] **Paso 5: Commit**

```bash
git add src/randomizer.py tests/test_randomizer.py
git commit -m "feat: implementar randomizador de preguntas y alternativas"
```

---

### Tarea 5: Generador de documentos .docx

**Archivos:**
- Crear: `src/doc_builder.py`
- Crear: `tests/test_doc_builder.py`

- [ ] **Paso 1: Escribir el test del generador de documentos**

```python
# tests/test_doc_builder.py
import os
import tempfile
from docx import Document
from src.models import Alternative, Question, Exam
from src.doc_builder import build_docx


def _make_exam() -> Exam:
    """Crea un examen de prueba."""
    questions = []
    for i in range(1, 4):
        alts = [
            Alternative(letter="a", text=f"Opción A de P{i}", image=None),
            Alternative(letter="b", text=f"Opción B de P{i}", image=None),
            Alternative(letter="c", text=f"Opción C de P{i}", image=None),
        ]
        questions.append(
            Question(
                number=i,
                text=f"Enunciado de la pregunta {i}",
                alternatives=alts,
                images=[],
            )
        )
    return Exam(
        title="Prueba de Ondas",
        questions=questions,
        header_paragraphs=["Prueba de Ondas", "Nombre: ________________"],
    )


def test_build_docx_creates_file():
    exam = _make_exam()
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        output_path = f.name
    try:
        build_docx(exam, output_path)
        assert os.path.exists(output_path)
        assert os.path.getsize(output_path) > 0
    finally:
        os.unlink(output_path)


def test_build_docx_contains_all_questions():
    exam = _make_exam()
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        output_path = f.name
    try:
        build_docx(exam, output_path)
        doc = Document(output_path)
        full_text = "\n".join(p.text for p in doc.paragraphs)

        # Debe contener las 3 preguntas
        assert "1-" in full_text
        assert "2-" in full_text
        assert "3-" in full_text

        # Debe contener alternativas
        assert "a)" in full_text
        assert "b)" in full_text
        assert "c)" in full_text

        # Debe contener el header
        assert "Prueba de Ondas" in full_text
        assert "Nombre:" in full_text
    finally:
        os.unlink(output_path)


def test_build_docx_preserves_question_format():
    exam = _make_exam()
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        output_path = f.name
    try:
        build_docx(exam, output_path)
        doc = Document(output_path)
        texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

        # Buscar patrón "N- enunciado"
        question_lines = [t for t in texts if t and t[0].isdigit() and "-" in t[:4]]
        assert len(question_lines) == 3
    finally:
        os.unlink(output_path)
```

- [ ] **Paso 2: Correr el test para verificar que falla**

Comando: `python -m pytest tests/test_doc_builder.py -v`
Esperado: FAIL con "ImportError: cannot import name 'build_docx'"

- [ ] **Paso 3: Implementar el generador de documentos**

```python
# src/doc_builder.py
"""Generador de archivos .docx a partir de un examen.

Reconstruye un documento Word con el formato esperado:
- Header (instrucciones, título)
- Preguntas con formato "N- enunciado"
- Alternativas con formato "x) texto"
- Imágenes embebidas cuando existen
"""
import io
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from src.models import Exam, Question


def _add_image_to_paragraph(paragraph, image_data: tuple[bytes, str]):
    """Agrega una imagen a un párrafo.

    Args:
        paragraph: Párrafo de python-docx donde insertar la imagen.
        image_data: Tupla (bytes, content_type) de la imagen.
    """
    image_bytes, content_type = image_data
    image_stream = io.BytesIO(image_bytes)
    run = paragraph.add_run()
    run.add_picture(image_stream, width=Inches(4.0))


def build_docx(exam: Exam, output_path: str) -> None:
    """Construye un archivo .docx a partir de un examen.

    El documento generado sigue el formato:
    1. Párrafos del header (instrucciones, título, etc.)
    2. Para cada pregunta:
       - "N- enunciado" (con imágenes si las hay)
       - "a) texto alternativa"
       - "b) texto alternativa"
       - etc.
    3. Línea vacía entre preguntas

    Args:
        exam: El examen a convertir en documento.
        output_path: Ruta donde guardar el archivo .docx generado.
    """
    doc = Document()

    # Escribir header (instrucciones, título, nombre, etc.)
    for header_text in exam.header_paragraphs:
        doc.add_paragraph(header_text)

    # Línea vacía entre header y preguntas
    if exam.header_paragraphs:
        doc.add_paragraph("")

    # Escribir cada pregunta
    for i, question in enumerate(exam.questions):
        # Enunciado de la pregunta: "N- texto"
        # Si el enunciado tiene múltiples líneas, la primera va con el número
        text_lines = question.text.split("\n")
        q_paragraph = doc.add_paragraph(f"{question.number}- {text_lines[0]}")

        # Líneas adicionales del enunciado (multilinea)
        for extra_line in text_lines[1:]:
            doc.add_paragraph(extra_line)

        # Imágenes del enunciado
        for image_data in question.images:
            img_paragraph = doc.add_paragraph()
            _add_image_to_paragraph(img_paragraph, image_data)

        # Alternativas
        for alt in question.alternatives:
            alt_text = f"{alt.letter}) {alt.text}"
            alt_paragraph = doc.add_paragraph(alt_text)

            # Imagen de la alternativa (si existe)
            if alt.image:
                _add_image_to_paragraph(alt_paragraph, alt.image)

        # Línea vacía entre preguntas (excepto después de la última)
        if i < len(exam.questions) - 1:
            doc.add_paragraph("")

    doc.save(output_path)
```

- [ ] **Paso 4: Correr el test para verificar que pasa**

Comando: `python -m pytest tests/test_doc_builder.py -v`
Esperado: PASS (3 tests)

- [ ] **Paso 5: Commit**

```bash
git add src/doc_builder.py tests/test_doc_builder.py
git commit -m "feat: implementar generador de documentos .docx"
```

---

### Tarea 6: Cliente de Google Drive

**Archivos:**
- Crear: `src/drive_client.py`

- [ ] **Paso 1: Implementar el cliente de Google Drive**

```python
# src/drive_client.py
"""Cliente para interactuar con Google Drive API.

Maneja autenticación OAuth2, descarga de archivos .docx y subida de archivos generados.
Usa credenciales almacenadas en credentials/client_secret.json y guarda el token
en credentials/token.json para reutilizar la sesión.
"""
import os
import io
import tempfile
from pathlib import Path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

# Scope necesario: lectura y escritura de archivos en Drive
SCOPES = ["https://www.googleapis.com/auth/drive"]

# Rutas de credenciales (relativas al directorio raíz del proyecto)
CREDENTIALS_DIR = Path(__file__).parent.parent / "credentials"
CLIENT_SECRET_PATH = CREDENTIALS_DIR / "client_secret.json"
TOKEN_PATH = CREDENTIALS_DIR / "token.json"


def _authenticate() -> Credentials:
    """Autentica con Google y devuelve credenciales válidas.

    Si existe un token guardado y es válido, lo reutiliza.
    Si el token expiró, lo refresca.
    Si no existe token, abre el navegador para login OAuth2.

    Returns:
        Credenciales de Google autenticadas.

    Raises:
        FileNotFoundError: Si no existe client_secret.json.
    """
    creds = None

    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CLIENT_SECRET_PATH.exists():
                raise FileNotFoundError(
                    f"No se encontró {CLIENT_SECRET_PATH}. "
                    f"Seguí las instrucciones en credentials/README.md para obtener las credenciales."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CLIENT_SECRET_PATH), SCOPES
            )
            creds = flow.run_local_server(port=0)

        # Guardar token para la próxima vez
        TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(TOKEN_PATH, "w") as token_file:
            token_file.write(creds.to_json())

    return creds


def _get_service():
    """Obtiene una instancia del servicio de Google Drive API.

    Returns:
        Objeto service de Google Drive API v3.
    """
    creds = _authenticate()
    return build("drive", "v3", credentials=creds)


def _resolve_path(service, path: str) -> str:
    """Resuelve una ruta tipo 'Carpeta/Subcarpeta/archivo.docx' a un file ID de Drive.

    Navega carpeta por carpeta buscando por nombre hasta llegar al archivo final.

    Args:
        service: Servicio de Google Drive API.
        path: Ruta al archivo en Drive (ej: "Guías/Ondas 2026. Ev. 1 F1.docx").

    Returns:
        El file ID del archivo en Google Drive.

    Raises:
        FileNotFoundError: Si no se encuentra algún componente de la ruta.
    """
    parts = path.strip("/").split("/")
    parent_id = "root"

    for i, part in enumerate(parts):
        is_last = i == len(parts) - 1
        query = f"name = '{part}' and '{parent_id}' in parents and trashed = false"

        if not is_last:
            # Es una carpeta
            query += " and mimeType = 'application/vnd.google-apps.folder'"

        results = (
            service.files()
            .list(q=query, fields="files(id, name, mimeType)", spaces="drive")
            .execute()
        )
        files = results.get("files", [])

        if not files:
            raise FileNotFoundError(
                f"No se encontró '{part}' en la ruta '{path}'. "
                f"Verificá que el archivo exista en Google Drive."
            )

        parent_id = files[0]["id"]

    return parent_id


def _get_parent_folder_id(service, path: str) -> str:
    """Obtiene el ID de la carpeta padre de un archivo.

    Args:
        service: Servicio de Google Drive API.
        path: Ruta completa al archivo.

    Returns:
        ID de la carpeta padre.
    """
    parts = path.strip("/").split("/")
    if len(parts) <= 1:
        return "root"

    parent_path = "/".join(parts[:-1])
    return _resolve_path(service, parent_path)


def download_docx(drive_path: str) -> str:
    """Descarga un archivo .docx desde Google Drive a un archivo temporal local.

    Args:
        drive_path: Ruta al archivo en Drive (ej: "Guías/Ondas 2026. Ev. 1 F1.docx").

    Returns:
        Ruta al archivo temporal descargado.

    Raises:
        FileNotFoundError: Si el archivo no existe en Drive.
    """
    service = _get_service()
    file_id = _resolve_path(service, drive_path)

    # Descargar el archivo
    request = service.files().get_media(fileId=file_id)
    temp_file = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)

    downloader = MediaIoBaseDownload(temp_file, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    temp_file.close()
    print(f"✅ Archivo descargado: {drive_path}")
    return temp_file.name


def upload_docx(local_path: str, drive_folder_path: str, filename: str) -> str:
    """Sube un archivo .docx a una carpeta específica en Google Drive.

    Args:
        local_path: Ruta al archivo .docx local a subir.
        drive_folder_path: Ruta de la carpeta destino en Drive (ej: "Guías").
        filename: Nombre del archivo en Drive (ej: "Ondas 2026. Ev. 1 F1 Forma A.docx").

    Returns:
        ID del archivo subido en Drive.
    """
    service = _get_service()

    if drive_folder_path:
        parent_id = _resolve_path(service, drive_folder_path)
    else:
        parent_id = "root"

    file_metadata = {"name": filename, "parents": [parent_id]}
    media = MediaFileUpload(
        local_path,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    file = (
        service.files()
        .create(body=file_metadata, media_body=media, fields="id")
        .execute()
    )

    print(f"✅ Subido: {filename}")
    return file.get("id")
```

- [ ] **Paso 2: Commit**

```bash
git add src/drive_client.py
git commit -m "feat: implementar cliente de Google Drive (auth, descarga, subida)"
```

> **Nota:** Los tests de este módulo requieren credenciales reales de Google Drive, por lo que se testean de forma integrada al correr el script principal. Un test automático requeriría mocks extensivos que salen del scope de este plan.

---

### Tarea 7: Script principal `main.py`

**Archivos:**
- Crear: `src/main.py`

- [ ] **Paso 1: Implementar el script principal**

```python
# src/main.py
"""Script principal del randomizador de pruebas.

Uso:
    python -m src.main <ruta_en_drive> <n_formas>

Ejemplo:
    python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 3

Esto genera 3 formas (A, B, C) del examen y las sube a la misma carpeta en Drive.
"""
import os
import sys
import tempfile
from pathlib import Path
from src.drive_client import download_docx, upload_docx
from src.parser import parse_docx
from src.randomizer import generate_forms
from src.doc_builder import build_docx


def get_form_suffix(index: int) -> str:
    """Genera el sufijo de la forma: A, B, C, ..., Z, AA, AB, ...

    Args:
        index: Índice de la forma (0-based).

    Returns:
        Sufijo string (ej: "A", "B", ..., "Z", "AA", "AB", ...).
    """
    if index < 26:
        return chr(ord("A") + index)
    else:
        # Para más de 26 formas: AA, AB, AC, ...
        first = chr(ord("A") + (index // 26) - 1)
        second = chr(ord("A") + (index % 26))
        return first + second


def main():
    """Punto de entrada principal del programa."""
    if len(sys.argv) < 3:
        print("Uso: python -m src.main <ruta_en_drive> <n_formas>")
        print('Ejemplo: python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 3')
        sys.exit(1)

    drive_path = sys.argv[1]
    n_forms = int(sys.argv[2])

    if n_forms < 1:
        print("❌ El número de formas debe ser al menos 1.")
        sys.exit(1)

    # Extraer nombre del archivo y carpeta en Drive
    path_parts = drive_path.strip("/").split("/")
    filename = path_parts[-1]  # "Ondas 2026. Ev. 1 F1.docx"
    drive_folder = "/".join(path_parts[:-1]) if len(path_parts) > 1 else ""
    base_name = filename.rsplit(".docx", 1)[0]  # "Ondas 2026. Ev. 1 F1"

    print(f"📄 Archivo: {drive_path}")
    print(f"📁 Carpeta destino: {drive_folder or 'raíz de Drive'}")
    print(f"🔄 Generando {n_forms} forma(s)...\n")

    # 1. Descargar archivo desde Google Drive
    print("⬇️  Descargando archivo de Google Drive...")
    local_path = download_docx(drive_path)

    try:
        # 2. Parsear el documento
        print("📖 Parseando documento...")
        exam = parse_docx(local_path)
        print(
            f"   Encontradas {len(exam.questions)} preguntas, "
            f"header con {len(exam.header_paragraphs)} línea(s).\n"
        )

        # 3. Generar N formas aleatorias
        print("🎲 Generando formas aleatorias...")
        forms = generate_forms(exam, n_forms)

        # 4. Construir y subir cada forma
        temp_dir = tempfile.mkdtemp()
        for i, form in enumerate(forms):
            suffix = get_form_suffix(i)
            form_filename = f"{base_name} Forma {suffix}.docx"
            form_path = os.path.join(temp_dir, form_filename)

            print(f"\n📝 Generando: {form_filename}")
            build_docx(form, form_path)

            print(f"⬆️  Subiendo a Drive...")
            upload_docx(form_path, drive_folder, form_filename)

            # Limpiar archivo temporal
            os.unlink(form_path)

        os.rmdir(temp_dir)

    finally:
        # Limpiar archivo descargado
        if os.path.exists(local_path):
            os.unlink(local_path)

    print(f"\n✅ ¡Listo! Se generaron {n_forms} forma(s) exitosamente.")
    print("   Revisá tu carpeta en Google Drive.")


if __name__ == "__main__":
    main()
```

- [ ] **Paso 2: Verificar ejecución sin argumentos**

Comando: `python -m src.main`
Esperado: Mensaje de uso y salida con código 1.

- [ ] **Paso 3: Commit**

```bash
git add src/main.py
git commit -m "feat: implementar script principal con flujo completo"
```

---

### Tarea 8: Test de integración end-to-end (sin Drive)

**Archivos:**
- Crear: `tests/test_integration.py`

- [ ] **Paso 1: Escribir test de integración local**

```python
# tests/test_integration.py
"""Test de integración que verifica el flujo completo sin Google Drive.

Crea un .docx de prueba, lo parsea, genera formas aleatorias,
y verifica que los archivos generados son válidos.
"""
import os
import tempfile
from docx import Document
from src.parser import parse_docx
from src.randomizer import generate_forms
from src.doc_builder import build_docx


def _create_full_exam_docx(path: str):
    """Crea un examen completo de prueba con 5 preguntas."""
    doc = Document()
    doc.add_paragraph("Evaluación de Física - Ondas y Sonido")
    doc.add_paragraph("Nombre: ________________  Fecha: ________")
    doc.add_paragraph("Instrucciones: Seleccione la alternativa correcta.")
    doc.add_paragraph("")

    questions = [
        (
            "¿Cuál es la velocidad del sonido en el aire a 20°C?",
            ["343 m/s", "300 m/s", "400 m/s", "250 m/s", "500 m/s"],
        ),
        (
            "¿Qué propiedad de una onda determina su tono?",
            ["Frecuencia", "Amplitud", "Longitud de onda", "Velocidad", "Fase"],
        ),
        (
            "¿Cuál es la unidad de medida de la intensidad sonora?",
            ["Decibel", "Hertz", "Metro", "Pascal", "Watt"],
        ),
        (
            "¿Qué fenómeno ocurre cuando dos ondas se superponen?",
            ["Interferencia", "Refracción", "Difracción", "Absorción", "Transmisión"],
        ),
        (
            "¿Qué tipo de onda necesita un medio material para propagarse?",
            ["Mecánica", "Electromagnética", "Gravitacional", "Nuclear", "Cuántica"],
        ),
    ]

    for i, (text, alts) in enumerate(questions, 1):
        doc.add_paragraph(f"{i}- {text}")
        for j, alt_text in enumerate(alts):
            letter = chr(ord("a") + j)
            doc.add_paragraph(f"{letter}) {alt_text}")
        doc.add_paragraph("")

    doc.save(path)


def test_full_pipeline_local():
    """Test completo: crear docx → parsear → randomizar → generar docx."""
    with tempfile.TemporaryDirectory() as tmpdir:
        # 1. Crear examen de prueba
        source_path = os.path.join(tmpdir, "examen_original.docx")
        _create_full_exam_docx(source_path)

        # 2. Parsear
        exam = parse_docx(source_path)
        assert len(exam.questions) == 5
        assert len(exam.header_paragraphs) >= 1

        # 3. Generar 3 formas
        forms = generate_forms(exam, n=3)
        assert len(forms) == 3

        # 4. Construir archivos .docx
        output_paths = []
        for i, form in enumerate(forms):
            suffix = chr(ord("A") + i)
            out_path = os.path.join(tmpdir, f"examen_original Forma {suffix}.docx")
            build_docx(form, out_path)
            output_paths.append(out_path)

        # 5. Verificar que los archivos existen y son válidos
        for out_path in output_paths:
            assert os.path.exists(out_path)
            assert os.path.getsize(out_path) > 0

            # Re-parsear para verificar estructura
            re_parsed = parse_docx(out_path)
            assert len(re_parsed.questions) == 5
            for q in re_parsed.questions:
                assert len(q.alternatives) == 5

        # 6. Verificar que las formas son distintas entre sí
        form_orders = []
        for out_path in output_paths:
            parsed = parse_docx(out_path)
            order = tuple(q.text for q in parsed.questions)
            form_orders.append(order)

        # Con 3 formas de 5 preguntas, es muy probable que al menos 2 sean distintas
        unique_orders = set(form_orders)
        assert len(unique_orders) >= 2, (
            "Las 3 formas tienen exactamente el mismo orden de preguntas — "
            "algo anda mal con el randomizador"
        )


def test_roundtrip_preserves_content():
    """Verifica que el contenido se preserva tras parse → build → parse."""
    with tempfile.TemporaryDirectory() as tmpdir:
        source_path = os.path.join(tmpdir, "original.docx")
        _create_full_exam_docx(source_path)

        # Parse original
        exam = parse_docx(source_path)
        original_texts = {q.text for q in exam.questions}
        original_alt_counts = {len(q.alternatives) for q in exam.questions}

        # Build nuevo documento (sin randomizar, usando el examen tal cual)
        rebuilt_path = os.path.join(tmpdir, "rebuilt.docx")
        build_docx(exam, rebuilt_path)

        # Re-parse
        rebuilt_exam = parse_docx(rebuilt_path)
        rebuilt_texts = {q.text for q in rebuilt_exam.questions}

        assert original_texts == rebuilt_texts
        assert all(len(q.alternatives) == 5 for q in rebuilt_exam.questions)
```

- [ ] **Paso 2: Correr el test de integración**

Comando: `python -m pytest tests/test_integration.py -v`
Esperado: PASS (2 tests)

- [ ] **Paso 3: Correr TODOS los tests**

Comando: `python -m pytest tests/ -v`
Esperado: PASS (todos los tests — ~16 tests en total)

- [ ] **Paso 4: Commit**

```bash
git add tests/test_integration.py
git commit -m "test: agregar test de integración end-to-end"
```

---

### Tarea 9: Documentación y README

**Archivos:**
- Crear: `README.md`

- [ ] **Paso 1: Crear README.md**

```markdown
# 🎲 Randomizador de Pruebas

Genera múltiples formas de una prueba de alternativas reordenando preguntas y opciones de respuesta.

## Requisitos

- Python 3.10+
- Credenciales de Google Drive API (ver `credentials/README.md`)

## Instalación

1. Clonar el repositorio
2. Crear entorno virtual e instalar dependencias:
   ```
   python -m venv venv
   source venv/bin/activate  # macOS/Linux
   pip install -r requirements.txt
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
python -m pytest tests/ -v
```
```

- [ ] **Paso 2: Commit**

```bash
git add README.md
git commit -m "docs: agregar README con instrucciones de uso"
```

---

## Resumen de tareas

| Tarea | Descripción | Archivos principales |
|-------|-------------|---------------------|
| 1 | Setup del proyecto | `requirements.txt`, `.gitignore` |
| 2 | Modelos de datos | `src/models.py` |
| 3 | Parser de .docx | `src/parser.py` |
| 4 | Randomizador | `src/randomizer.py` |
| 5 | Generador de .docx | `src/doc_builder.py` |
| 6 | Cliente Google Drive | `src/drive_client.py` |
| 7 | Script principal | `src/main.py` |
| 8 | Test de integración | `tests/test_integration.py` |
| 9 | README | `README.md` |