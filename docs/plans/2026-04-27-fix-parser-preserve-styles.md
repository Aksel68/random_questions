# Fix Parser & Preserve Styles — Plan de Implementación

> **Para agentes:** Usá `@rules/skills/subagents.md` (recomendado) o `@rules/skills/executing-plans.md` para implementar este plan tarea por tarea. Los pasos usan sintaxis checkbox (`- [ ]`) para tracking.

**Goal:** Reescribir el parser y doc_builder para que (1) extraigan preguntas desde celdas de una tabla dentro del .docx, y (2) preserven todos los estilos, imágenes y estructura del documento original al generar las formas randomizadas.

**Arquitectura:** En lugar de recrear el documento desde cero, el nuevo enfoque **clona el .docx original** y manipula directamente los elementos XML de las celdas de la tabla de preguntas. El parser extrae "bloques de pregunta" como listas de objetos `Paragraph` XML (preservando runs, formato e imágenes). El randomizador reordena estos bloques. El builder clona el documento N veces y reemplaza el contenido de cada celda con el bloque correspondiente.

**Tech Stack:** Python 3.10+, `python-docx` (manipulación de .docx), `copy.deepcopy` + `lxml` (clonación de elementos XML), `random` (stdlib)

---

## Estructura del documento analizado

```
Body:
├── Párrafos superiores: Título "Prueba 1. Ondas..." + "Forma 1"
├── [sdt element]: Logo/encabezado del colegio
├── TABLE[0] (2x4): Datos alumno (NOMBRE, CURSO, FECHA, PUNTAJE)
├── TABLE[1] (1x1): Instrucciones
├── Párrafo: "Ítem I. Marque la alternativa..."
├── TABLE[2] (18x2): PREGUNTAS — 2 preguntas por fila, 36 total
├── Párrafos finales: "¡Has llegado al final..." + imagen final
```

Cada celda de TABLE[2] contiene:
- Primer párrafo: enunciado con formato `N-` o `N.` (puede tener imágenes inline)
- Párrafos intermedios: continuación del enunciado (multilinea) o imágenes sueltas
- Párrafos de alternativas: `a)`, `b)`, `c)`, `d)` (pueden tener imágenes)
- Párrafos vacíos al final (padding)

## Estructura de archivos

| Archivo | Responsabilidad |
|---|---|
| `src/models.py` | **Reescribir** — Nuevos modelos que almacenan elementos XML en lugar de texto plano |
| `src/parser.py` | **Reescribir** — Extrae preguntas desde celdas de tabla, preservando XML |
| `src/randomizer.py` | **Modificar** — Trabaja con los nuevos modelos (listas de elementos XML) |
| `src/doc_builder.py` | **Reescribir** — Clona el documento original y reemplaza contenido de celdas |
| `src/main.py` | **Modificar menor** — Pasa ruta del archivo original al builder |
| `tests/test_parser.py` | **Reescribir** — Tests adaptados a tabla |
| `tests/test_randomizer.py` | **Modificar** — Adaptar a nuevos modelos |
| `tests/test_doc_builder.py` | **Reescribir** — Tests de preservación de formato |
| `tests/test_integration.py` | **Reescribir** — Test end-to-end con tabla |

---

### Tarea 1: Rediseñar modelos de datos

**Archivos:**
- Modificar: `src/models.py`

- [ ] **Paso 1: Reescribir `src/models.py`**

```python
# src/models.py
"""Modelos de datos para el randomizador de pruebas.

Los modelos almacenan los elementos XML de python-docx directamente,
lo que permite preservar todo el formato (negrita, cursiva, imágenes, etc.)
al reconstruir el documento.
"""
from dataclasses import dataclass, field
from typing import Optional
from copy import deepcopy
from lxml import etree


@dataclass
class Alternative:
    """Una alternativa de una pregunta de selección múltiple.

    Attributes:
        letter: Letra de la alternativa (a, b, c, d...).
        paragraphs: Lista de elementos XML <w:p> que componen la alternativa.
            El primer párrafo contiene "x) texto" con formato preservado.
    """
    letter: str
    paragraphs: list  # list of lxml elements (<w:p>)

    def clone(self) -> "Alternative":
        """Crea una copia profunda de la alternativa con XML clonado."""
        return Alternative(
            letter=self.letter,
            paragraphs=[deepcopy(p) for p in self.paragraphs],
        )

    @property
    def text(self) -> str:
        """Texto plano de la alternativa (para debugging/testing)."""
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        parts = []
        for p in self.paragraphs:
            runs = p.findall(".//w:r/w:t", nsmap)
            parts.append("".join(r.text or "" for r in runs))
        return " ".join(parts).strip()


@dataclass
class Question:
    """Una pregunta completa con enunciado y alternativas.

    Attributes:
        number: Número original de la pregunta.
        enunciado_paragraphs: Elementos XML <w:p> del enunciado (antes de las alternativas).
        alternatives: Lista de alternativas.
        trailing_empty_paragraphs: Párrafos vacíos al final de la celda (padding).
    """
    number: int
    enunciado_paragraphs: list  # list of lxml elements (<w:p>)
    alternatives: list[Alternative] = field(default_factory=list)
    trailing_empty_paragraphs: list = field(default_factory=list)  # list of lxml elements

    def clone(self) -> "Question":
        """Crea una copia profunda de la pregunta con XML clonado."""
        return Question(
            number=self.number,
            enunciado_paragraphs=[deepcopy(p) for p in self.enunciado_paragraphs],
            alternatives=[alt.clone() for alt in self.alternatives],
            trailing_empty_paragraphs=[deepcopy(p) for p in self.trailing_empty_paragraphs],
        )

    @property
    def text(self) -> str:
        """Texto plano del enunciado (para debugging/testing)."""
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        parts = []
        for p in self.enunciado_paragraphs:
            runs = p.findall(".//w:r/w:t", nsmap)
            parts.append("".join(r.text or "" for r in runs))
        return " ".join(parts).strip()


@dataclass
class Exam:
    """Un examen completo parseado de un .docx con tabla de preguntas.

    Attributes:
        source_path: Ruta al archivo .docx original (para clonarlo al generar formas).
        questions: Lista de preguntas extraídas de la tabla.
        questions_table_index: Índice de la tabla que contiene las preguntas.
        questions_per_row: Número de preguntas por fila en la tabla (normalmente 2).
    """
    source_path: str
    questions: list[Question] = field(default_factory=list)
    questions_table_index: int = 2
    questions_per_row: int = 2
```

- [ ] **Paso 2: Verificar que no hay errores de sintaxis**

Comando: `uv run python -c "from src.models import Alternative, Question, Exam; print('OK')"`
Esperado: `OK`

- [ ] **Paso 3: Commit**

```bash
git add src/models.py
git commit -m "refactor: rewrite models to store XML elements for style preservation"
```

---

### Tarea 2: Reescribir el parser para extraer desde tablas

**Archivos:**
- Modificar: `src/parser.py`
- Modificar: `tests/test_parser.py`

- [ ] **Paso 1: Reescribir `src/parser.py`**

```python
# src/parser.py
"""Parser de archivos .docx con formato de pruebas de alternativas en tabla.

Estructura esperada del documento:
- Tablas iniciales con datos del alumno e instrucciones
- Una tabla principal (normalmente la última/más grande) con preguntas
- Cada celda de la tabla de preguntas contiene:
  - Enunciado: párrafo(s) que empiezan con "N-" o "N."
  - Alternativas: párrafos que empiezan con "x)" donde x es letra
  - Párrafos vacíos de padding (al final)

El parser extrae directamente los elementos XML <w:p> para preservar
todo el formato (negrita, cursiva, imágenes, colores, fuentes).
"""
import re
from copy import deepcopy
from docx import Document
from lxml import etree
from src.models import Alternative, Question, Exam

# Namespace de Word XML
NSMAP = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

# Regex para detectar inicio de pregunta: "N-" o "N." donde N es uno o más dígitos
QUESTION_PATTERN = re.compile(r"^(\d+)\s*[-–—.]\s*")

# Regex para detectar inicio de alternativa: "x)" donde x es una letra
ALTERNATIVE_PATTERN = re.compile(r"^([a-zA-Z])\)\s*")


def _get_paragraph_text(p_element) -> str:
    """Extrae texto plano de un elemento <w:p> XML.

    Args:
        p_element: Elemento lxml representando un párrafo de Word.

    Returns:
        Texto concatenado de todos los runs del párrafo.
    """
    runs = p_element.findall(".//w:r/w:t", NSMAP)
    return "".join(r.text or "" for r in runs).strip()


def _find_questions_table(doc: Document) -> tuple[int, "Table"]:
    """Encuentra la tabla que contiene las preguntas del examen.

    Heurística: es la tabla con más filas (la tabla de preguntas siempre
    tiene muchas más filas que las tablas de datos/instrucciones).

    Args:
        doc: Documento python-docx.

    Returns:
        Tupla (índice_tabla, objeto_tabla).

    Raises:
        ValueError: Si no se encuentran tablas en el documento.
    """
    if not doc.tables:
        raise ValueError("El documento no contiene tablas.")

    max_rows = 0
    best_idx = 0
    for i, table in enumerate(doc.tables):
        if len(table.rows) > max_rows:
            max_rows = len(table.rows)
            best_idx = i

    return best_idx, doc.tables[best_idx]


def _parse_cell(cell) -> Question | None:
    """Parsea una celda de la tabla de preguntas.

    Extrae el enunciado, alternativas y párrafos vacíos de trailing.
    Preserva los elementos XML originales (deep copy) para mantener formato.

    Args:
        cell: Objeto Cell de python-docx.

    Returns:
        Un objeto Question si la celda contiene una pregunta, None si está vacía.
    """
    # Acceder a los elementos XML <w:p> dentro del <w:tc>
    tc_element = cell._tc
    p_elements = tc_element.findall("w:p", NSMAP)

    if not p_elements:
        return None

    # Verificar que la celda contiene una pregunta
    first_text = _get_paragraph_text(p_elements[0])
    q_match = QUESTION_PATTERN.match(first_text)
    if not q_match:
        return None

    question_number = int(q_match.group(1))

    # Clasificar párrafos: enunciado, alternativas, trailing empty
    enunciado_paragraphs = []
    alternatives = []
    trailing_empty = []
    current_alt_paragraphs = []
    current_alt_letter = None
    in_alternatives = False

    for p_elem in p_elements:
        text = _get_paragraph_text(p_elem)
        alt_match = ALTERNATIVE_PATTERN.match(text)

        if alt_match:
            # Si había una alternativa en curso, guardarla
            if current_alt_letter is not None:
                alternatives.append(Alternative(
                    letter=current_alt_letter,
                    paragraphs=[deepcopy(pp) for pp in current_alt_paragraphs],
                ))

            in_alternatives = True
            current_alt_letter = alt_match.group(1).lower()
            current_alt_paragraphs = [p_elem]
            # Reset trailing (previous empty paragraphs were part of content)
            if trailing_empty:
                # Were these actually between alternatives? Add to previous alt
                if alternatives:
                    alternatives[-1].paragraphs.extend(
                        [deepcopy(pp) for pp in trailing_empty]
                    )
                trailing_empty = []

        elif not text and in_alternatives:
            # Párrafo vacío después de las alternativas — es trailing
            trailing_empty.append(p_elem)

        elif not text and not in_alternatives:
            # Párrafo vacío en el enunciado — puede ser separador antes de imagen
            enunciado_paragraphs.append(p_elem)

        else:
            if not in_alternatives:
                enunciado_paragraphs.append(p_elem)
            else:
                # Texto no-alternativa después de alternativas — raro, tratar como trailing
                # o parte de la última alternativa (e.g. continuación multilinea)
                current_alt_paragraphs.append(p_elem)
                trailing_empty = []

    # Guardar la última alternativa
    if current_alt_letter is not None:
        alternatives.append(Alternative(
            letter=current_alt_letter,
            paragraphs=[deepcopy(pp) for pp in current_alt_paragraphs],
        ))

    if not alternatives:
        return None

    return Question(
        number=question_number,
        enunciado_paragraphs=[deepcopy(p) for p in enunciado_paragraphs],
        alternatives=alternatives,
        trailing_empty_paragraphs=[deepcopy(p) for p in trailing_empty],
    )


def parse_docx(file_path: str) -> Exam:
    """Parsea un archivo .docx de prueba de alternativas con preguntas en tabla.

    Args:
        file_path: Ruta al archivo .docx.

    Returns:
        Un objeto Exam con las preguntas extraídas, preservando elementos XML.
    """
    doc = Document(file_path)

    table_idx, questions_table = _find_questions_table(doc)

    questions: list[Question] = []

    for row in questions_table.rows:
        for cell in row.cells:
            question = _parse_cell(cell)
            if question is not None:
                questions.append(question)

    # Determinar preguntas por fila
    cols_count = len(questions_table.columns) if questions_table.columns else 2

    return Exam(
        source_path=file_path,
        questions=questions,
        questions_table_index=table_idx,
        questions_per_row=cols_count,
    )
```

- [ ] **Paso 2: Reescribir `tests/test_parser.py`**

```python
# tests/test_parser.py
"""Tests del parser v2 — extracción desde tablas con formato preservado."""
import os
import tempfile
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from src.parser import parse_docx
from src.models import Exam


def _create_table_exam_docx(path: str, num_questions: int = 6):
    """Crea un .docx con estructura de tabla similar al documento real.

    - Tabla 0 (1x2): Header (simulado)
    - Tabla 1 (Nx2): Preguntas en pares (2 por fila)
    """
    doc = Document()
    doc.add_paragraph("Prueba de Física")

    # Tabla de datos (header)
    header_table = doc.add_table(rows=1, cols=2)
    header_table.cell(0, 0).text = "NOMBRE"
    header_table.cell(0, 1).text = ""

    doc.add_paragraph("Ítem I. Marque la alternativa correcta.")

    # Tabla de preguntas (2 columnas)
    rows_needed = (num_questions + 1) // 2
    q_table = doc.add_table(rows=rows_needed, cols=2)

    questions_data = [
        ("¿Cuál es la velocidad de la luz?", ["3x10^8 m/s", "3x10^6 m/s", "3x10^10 m/s", "3x10^4 m/s"]),
        ("¿Qué tipo de onda es el sonido?", ["Transversal", "Longitudinal", "Electromagnética", "Estacionaria"]),
        ("¿Cuál es la unidad de frecuencia?", ["Metro", "Segundo", "Hertz", "Newton"]),
        ("¿Qué transmiten las ondas?", ["Materia", "Fuerza", "Energía", "Calor"]),
        ("¿Qué es un pulso?", ["Perturbación aislada", "Onda periódica", "Onda continua", "Vibración"]),
        ("¿Qué necesitan las ondas mecánicas?", ["Vacío", "Luz", "Medio material", "Electricidad"]),
    ]

    for i in range(min(num_questions, len(questions_data))):
        row_idx = i // 2
        col_idx = i % 2
        cell = q_table.cell(row_idx, col_idx)

        q_text, alts = questions_data[i]
        # Limpiar celda (viene con un párrafo vacío por defecto)
        cell.paragraphs[0].text = f"{i+1}- {q_text}"

        for j, alt_text in enumerate(alts):
            letter = chr(ord("a") + j)
            cell.add_paragraph(f"{letter}) {alt_text}")

    doc.save(path)


def test_parse_docx_extracts_questions_from_table():
    """Verifica que el parser extrae preguntas de la tabla correctamente."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name
    try:
        _create_table_exam_docx(path, num_questions=6)
        exam = parse_docx(path)

        assert isinstance(exam, Exam)
        assert len(exam.questions) == 6
        assert exam.questions_per_row == 2

        # Verificar primera pregunta
        q1 = exam.questions[0]
        assert q1.number == 1
        assert "velocidad de la luz" in q1.text.lower()
        assert len(q1.alternatives) == 4
        assert q1.alternatives[0].letter == "a"
    finally:
        os.unlink(path)


def test_parse_docx_preserves_formatting():
    """Verifica que el parser preserva los elementos XML con formato."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name
    try:
        doc = Document()
        q_table = doc.add_table(rows=1, cols=1)
        cell = q_table.cell(0, 0)

        # Pregunta con negrita
        p = cell.paragraphs[0]
        run1 = p.add_run("1- ¿Qué es la ")
        run2 = p.add_run("amplitud")
        run2.bold = True
        run3 = p.add_run("?")

        cell.add_paragraph("a) Distancia máxima")
        cell.add_paragraph("b) Frecuencia")
        cell.add_paragraph("c) Período")
        cell.add_paragraph("d) Velocidad")

        doc.save(path)

        exam = parse_docx(path)
        assert len(exam.questions) == 1
        q = exam.questions[0]
        assert q.number == 1
        assert len(q.alternatives) == 4

        # Verificar que el XML del enunciado tiene runs con formato
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        runs = q.enunciado_paragraphs[0].findall(".//w:r", nsmap)
        # Debe haber al menos 3 runs (texto + bold + texto)
        assert len(runs) >= 3
    finally:
        os.unlink(path)


def test_parse_docx_handles_dot_format():
    """Verifica que el parser detecta preguntas con formato 'N.' además de 'N-'."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name
    try:
        doc = Document()
        q_table = doc.add_table(rows=1, cols=1)
        cell = q_table.cell(0, 0)
        cell.paragraphs[0].text = "33. De los siguientes, ¿cuál NO es una onda?"
        cell.add_paragraph("a) Luz")
        cell.add_paragraph("b) Sonido")
        cell.add_paragraph("c) Vuelo de abeja")
        cell.add_paragraph("d) Terremoto")

        doc.save(path)

        exam = parse_docx(path)
        assert len(exam.questions) == 1
        assert exam.questions[0].number == 33
    finally:
        os.unlink(path)


def test_parse_docx_handles_multiline_enunciado():
    """Verifica que el enunciado multilinea se captura completo."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name
    try:
        doc = Document()
        q_table = doc.add_table(rows=1, cols=1)
        cell = q_table.cell(0, 0)
        cell.paragraphs[0].text = "3-\"Corresponde a la posición de reposo en el"
        cell.add_paragraph("que se encuentra un cuerpo antes de")
        cell.add_paragraph("empezar a vibrar\". Esta definición corresponde a:")
        cell.add_paragraph("a) Amplitud")
        cell.add_paragraph("b) Longitud de onda")
        cell.add_paragraph("c) Valle")
        cell.add_paragraph("d) Punto de equilibrio")

        doc.save(path)

        exam = parse_docx(path)
        assert len(exam.questions) == 1
        q = exam.questions[0]
        assert q.number == 3
        assert len(q.enunciado_paragraphs) == 3
        assert len(q.alternatives) == 4
    finally:
        os.unlink(path)


def test_parse_finds_correct_table():
    """Verifica que el parser identifica la tabla de preguntas (la más grande)."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name
    try:
        doc = Document()
        # Tabla pequeña (header)
        t1 = doc.add_table(rows=2, cols=4)
        t1.cell(0, 0).text = "NOMBRE"

        # Tabla pequeña (instrucciones)
        t2 = doc.add_table(rows=1, cols=1)
        t2.cell(0, 0).text = "INSTRUCCIONES"

        # Tabla grande (preguntas) — esta debe ser detectada
        t3 = doc.add_table(rows=5, cols=2)
        for i in range(10):
            row_idx = i // 2
            col_idx = i % 2
            cell = t3.cell(row_idx, col_idx)
            cell.paragraphs[0].text = f"{i+1}- Pregunta {i+1}"
            cell.add_paragraph("a) Opción A")
            cell.add_paragraph("b) Opción B")
            cell.add_paragraph("c) Opción C")
            cell.add_paragraph("d) Opción D")

        doc.save(path)

        exam = parse_docx(path)
        assert exam.questions_table_index == 2  # Tercera tabla (index 2)
        assert len(exam.questions) == 10
    finally:
        os.unlink(path)
```

- [ ] **Paso 3: Correr los tests**

Comando: `uv run pytest tests/test_parser.py -v`
Esperado: PASS (5 tests)

- [ ] **Paso 4: Commit**

```bash
git add src/parser.py tests/test_parser.py
git commit -m "refactor: rewrite parser to extract questions from table cells preserving XML"
```

---

### Tarea 3: Adaptar el randomizador a los nuevos modelos

**Archivos:**
- Modificar: `src/randomizer.py`
- Modificar: `tests/test_randomizer.py`

- [ ] **Paso 1: Reescribir `src/randomizer.py`**

```python
# src/randomizer.py
"""Generador de formas aleatorias a partir de un examen.

Produce N versiones del examen con preguntas y alternativas en orden aleatorio.
Trabaja con elementos XML clonados — preserva todo el formato original.
"""
import re
import random
from copy import deepcopy
from lxml import etree
from src.models import Alternative, Question, Exam

NSMAP = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

# Patrones para reemplazar número de pregunta y letra de alternativa en el XML
QUESTION_NUMBER_PATTERN = re.compile(r"^(\d+)\s*[-–—.]\s*")
ALTERNATIVE_LETTER_PATTERN = re.compile(r"^([a-zA-Z])\)\s*")


def _get_first_text_element(p_element):
    """Obtiene el primer elemento <w:t> con texto de un párrafo.

    Args:
        p_element: Elemento XML <w:p>.

    Returns:
        El primer <w:t> que contiene texto, o None.
    """
    for t_elem in p_element.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
        if t_elem.text:
            return t_elem
    return None


def _renumber_question(question: Question, new_number: int) -> Question:
    """Cambia el número de una pregunta en su XML.

    Busca el patrón "N-" o "N." en el primer <w:t> del primer párrafo
    del enunciado y lo reemplaza con el nuevo número.

    Args:
        question: La pregunta (ya clonada).
        new_number: El nuevo número a asignar.

    Returns:
        La misma pregunta con número actualizado.
    """
    question.number = new_number

    if not question.enunciado_paragraphs:
        return question

    first_p = question.enunciado_paragraphs[0]
    t_elem = _get_first_text_element(first_p)
    if t_elem is not None and t_elem.text:
        # Reemplazar "N-" o "N." con el nuevo número
        t_elem.text = QUESTION_NUMBER_PATTERN.sub(f"{new_number}- ", t_elem.text)

    return question


def _relabel_alternative(alt: Alternative, new_letter: str) -> Alternative:
    """Cambia la letra de una alternativa en su XML.

    Busca el patrón "x)" en el primer <w:t> del primer párrafo
    y lo reemplaza con la nueva letra.

    Args:
        alt: La alternativa (ya clonada).
        new_letter: La nueva letra (a, b, c, d...).

    Returns:
        La misma alternativa con letra actualizada.
    """
    alt.letter = new_letter

    if not alt.paragraphs:
        return alt

    first_p = alt.paragraphs[0]
    t_elem = _get_first_text_element(first_p)
    if t_elem is not None and t_elem.text:
        t_elem.text = ALTERNATIVE_LETTER_PATTERN.sub(f"{new_letter}) ", t_elem.text)

    return alt


def _shuffle_alternatives(alternatives: list[Alternative]) -> list[Alternative]:
    """Reordena las alternativas y reasigna letras.

    Args:
        alternatives: Lista de alternativas (ya clonadas).

    Returns:
        Lista de alternativas en orden aleatorio con letras reasignadas.
    """
    shuffled = list(alternatives)
    random.shuffle(shuffled)
    for i, alt in enumerate(shuffled):
        new_letter = chr(ord("a") + i)
        _relabel_alternative(alt, new_letter)
    return shuffled


def generate_forms(exam: Exam, n: int) -> list[list[Question]]:
    """Genera N listas de preguntas randomizadas.

    Cada forma es una lista de Questions con:
    - Orden de preguntas aleatorio
    - Números reasignados (1, 2, 3...)
    - Alternativas en orden aleatorio con letras reasignadas

    Args:
        exam: El examen original parseado.
        n: Número de formas a generar.

    Returns:
        Lista de N listas de preguntas randomizadas (con XML clonado).

    Raises:
        ValueError: Si n < 1.
    """
    if n < 1:
        raise ValueError("El número de formas debe ser al menos 1")

    forms: list[list[Question]] = []

    for _ in range(n):
        # Clonar todas las preguntas
        cloned_questions = [q.clone() for q in exam.questions]

        # Mezclar orden de preguntas
        random.shuffle(cloned_questions)

        # Renumerar y mezclar alternativas
        for i, q in enumerate(cloned_questions):
            _renumber_question(q, i + 1)
            q.alternatives = _shuffle_alternatives(q.alternatives)

        forms.append(cloned_questions)

    return forms
```

- [ ] **Paso 2: Reescribir `tests/test_randomizer.py`**

```python
# tests/test_randomizer.py
"""Tests del randomizador v2 — trabaja con elementos XML."""
import os
import tempfile
from docx import Document
from src.parser import parse_docx
from src.randomizer import generate_forms
from src.models import Exam


def _create_test_exam(num_questions=5) -> Exam:
    """Crea un examen de prueba con tabla de preguntas."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name

    doc = Document()
    rows_needed = (num_questions + 1) // 2
    q_table = doc.add_table(rows=rows_needed, cols=2)

    for i in range(num_questions):
        row_idx = i // 2
        col_idx = i % 2
        cell = q_table.cell(row_idx, col_idx)
        cell.paragraphs[0].text = f"{i+1}- Pregunta número {i+1} sobre ondas"
        cell.add_paragraph(f"a) Opción A de P{i+1}")
        cell.add_paragraph(f"b) Opción B de P{i+1}")
        cell.add_paragraph(f"c) Opción C de P{i+1}")
        cell.add_paragraph(f"d) Opción D de P{i+1}")

    doc.save(path)
    exam = parse_docx(path)
    os.unlink(path)
    return exam


def test_generate_forms_returns_correct_count():
    exam = _create_test_exam(5)
    forms = generate_forms(exam, n=3)
    assert len(forms) == 3


def test_generate_forms_preserves_all_questions():
    exam = _create_test_exam(5)
    forms = generate_forms(exam, n=2)
    for form in forms:
        assert len(form) == 5
        # Todas las preguntas originales deben estar presentes
        original_texts = {q.text for q in exam.questions}
        form_texts = {q.text for q in form}
        assert original_texts == form_texts


def test_generate_forms_renumbers_questions():
    exam = _create_test_exam(5)
    forms = generate_forms(exam, n=2)
    for form in forms:
        numbers = [q.number for q in form]
        assert numbers == [1, 2, 3, 4, 5]


def test_generate_forms_relabels_alternatives():
    exam = _create_test_exam(5)
    forms = generate_forms(exam, n=2)
    for form in forms:
        for q in form:
            letters = [a.letter for a in q.alternatives]
            assert letters == ["a", "b", "c", "d"]


def test_generate_forms_shuffles_order():
    """Con 10 formas de 5 preguntas, al menos una debe diferir del original."""
    exam = _create_test_exam(5)
    forms = generate_forms(exam, n=10)
    original_order = [q.text for q in exam.questions]
    any_different = any(
        [q.text for q in form] != original_order for form in forms
    )
    assert any_different


def test_generate_forms_shuffles_alternatives():
    """Al menos una alternativa en alguna forma debe estar reordenada."""
    exam = _create_test_exam(5)
    forms = generate_forms(exam, n=10)
    original_alt_texts = [a.text for a in exam.questions[0].alternatives]
    any_different = False
    for form in forms:
        for q in form:
            alt_texts = [a.text for a in q.alternatives]
            if alt_texts != original_alt_texts:
                any_different = True
                break
        if any_different:
            break
    assert any_different


def test_generate_forms_does_not_mutate_original():
    exam = _create_test_exam(5)
    original_texts = [q.text for q in exam.questions]
    original_numbers = [q.number for q in exam.questions]
    generate_forms(exam, n=5)
    assert [q.text for q in exam.questions] == original_texts
    assert [q.number for q in exam.questions] == original_numbers
```

- [ ] **Paso 3: Correr los tests**

Comando: `uv run pytest tests/test_randomizer.py -v`
Esperado: PASS (7 tests)

- [ ] **Paso 4: Commit**

```bash
git add src/randomizer.py tests/test_randomizer.py
git commit -m "refactor: adapt randomizer to work with XML-based question models"
```

---

### Tarea 4: Reescribir el doc_builder para clonar el documento original

**Archivos:**
- Modificar: `src/doc_builder.py`
- Modificar: `tests/test_doc_builder.py`

- [ ] **Paso 1: Reescribir `src/doc_builder.py`**

```python
# src/doc_builder.py
"""Generador de archivos .docx por clonación del documento original.

En lugar de recrear el documento desde cero, este módulo:
1. Copia el archivo .docx original byte a byte
2. Abre la copia con python-docx
3. Reemplaza el contenido de cada celda de la tabla de preguntas
   con las preguntas randomizadas (preservando XML/formato/imágenes)
4. Actualiza el texto "Forma X" en el encabezado

Esto garantiza que TODO el formato, estilos, imágenes, logos, tablas
de datos e instrucciones se preservan exactamente como en el original.
"""
import re
import shutil
from copy import deepcopy
from docx import Document
from lxml import etree
from src.models import Question

NSMAP = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def _clear_cell(cell):
    """Elimina todos los párrafos de una celda excepto las propiedades.

    Las celdas de Word DEBEN tener al menos un párrafo (según el estándar),
    pero nosotros las vaciamos y luego insertamos los nuevos párrafos.

    Args:
        cell: Objeto Cell de python-docx.
    """
    tc = cell._tc
    # Remover todos los <w:p> existentes
    for p in tc.findall("w:p", NSMAP):
        tc.remove(p)


def _add_paragraphs_to_cell(cell, paragraphs: list):
    """Agrega una lista de elementos <w:p> XML a una celda.

    Args:
        cell: Objeto Cell de python-docx.
        paragraphs: Lista de elementos lxml <w:p> a insertar.
    """
    tc = cell._tc
    for p_elem in paragraphs:
        tc.append(deepcopy(p_elem))


def _build_cell_content(question: Question) -> list:
    """Construye la lista de párrafos XML para una celda a partir de una pregunta.

    Orden:
    1. Párrafos del enunciado (con número, formato e imágenes)
    2. Párrafos de cada alternativa (con letra, formato e imágenes)
    3. Párrafos vacíos de trailing (padding original)

    Args:
        question: Pregunta randomizada con XML preservado.

    Returns:
        Lista de elementos <w:p> lxml.
    """
    paragraphs = []

    # Enunciado
    for p in question.enunciado_paragraphs:
        paragraphs.append(p)

    # Alternativas
    for alt in question.alternatives:
        for p in alt.paragraphs:
            paragraphs.append(p)

    # Trailing empty paragraphs (preservar layout/spacing)
    for p in question.trailing_empty_paragraphs:
        paragraphs.append(p)

    return paragraphs


def _update_form_label(doc: Document, form_suffix: str):
    """Busca y reemplaza el texto 'Forma X' en los párrafos superiores.

    Busca un párrafo que contenga la palabra "Forma" seguida de un identificador
    y lo reemplaza con el nuevo sufijo.

    Args:
        doc: Documento python-docx.
        form_suffix: Sufijo de la forma (ej: "A", "B", "C").
    """
    form_pattern = re.compile(r"[Ff]orma\s+\S+")

    for paragraph in doc.paragraphs:
        full_text = paragraph.text
        if form_pattern.search(full_text):
            # Reemplazar en los runs
            for run in paragraph.runs:
                if form_pattern.search(run.text):
                    run.text = form_pattern.sub(f"Forma {form_suffix}", run.text)
            break  # Solo reemplazar la primera ocurrencia


def build_docx(
    source_path: str,
    questions: list[Question],
    output_path: str,
    form_suffix: str,
    questions_table_index: int = 2,
    questions_per_row: int = 2,
) -> None:
    """Genera un .docx randomizado clonando el original y reemplazando preguntas.

    Args:
        source_path: Ruta al .docx original.
        questions: Lista de preguntas randomizadas (con XML preservado).
        output_path: Ruta donde guardar el nuevo .docx.
        form_suffix: Sufijo de la forma ("A", "B", etc.) para actualizar el label.
        questions_table_index: Índice de la tabla de preguntas en el documento.
        questions_per_row: Número de preguntas por fila (columnas de la tabla).
    """
    # 1. Copiar el archivo original
    shutil.copy2(source_path, output_path)

    # 2. Abrir la copia y modificar
    doc = Document(output_path)

    # 3. Actualizar label "Forma X"
    _update_form_label(doc, form_suffix)

    # 4. Obtener la tabla de preguntas
    questions_table = doc.tables[questions_table_index]

    # 5. Reemplazar contenido de cada celda
    q_index = 0
    for row in questions_table.rows:
        for col_idx in range(questions_per_row):
            if q_index >= len(questions):
                break
            cell = row.cells[col_idx]

            # Limpiar la celda
            _clear_cell(cell)

            # Insertar los párrafos de la pregunta randomizada
            cell_content = _build_cell_content(questions[q_index])
            _add_paragraphs_to_cell(cell, cell_content)

            q_index += 1

    # 6. Guardar
    doc.save(output_path)
```

- [ ] **Paso 2: Reescribir `tests/test_doc_builder.py`**

```python
# tests/test_doc_builder.py
"""Tests del doc_builder v2 — clonación con preservación de formato."""
import os
import tempfile
from docx import Document
from docx.shared import Pt
from src.parser import parse_docx
from src.randomizer import generate_forms
from src.doc_builder import build_docx


def _create_source_exam(path: str):
    """Crea un documento fuente con tabla de preguntas y formato."""
    doc = Document()

    # Título con formato
    p = doc.add_paragraph()
    run = p.add_run("Prueba de Física - Ondas")
    run.bold = True
    run.font.size = Pt(14)

    # Label de forma
    p2 = doc.add_paragraph()
    run2 = p2.add_run("Forma 1")
    run2.bold = True

    # Tabla header
    ht = doc.add_table(rows=1, cols=2)
    ht.cell(0, 0).text = "NOMBRE"
    ht.cell(0, 1).text = ""

    # Tabla instrucciones
    it = doc.add_table(rows=1, cols=1)
    it.cell(0, 0).text = "INSTRUCCIONES: Lea con atención"

    # Tabla preguntas (3 filas x 2 cols = 6 preguntas)
    qt = doc.add_table(rows=3, cols=2)
    for i in range(6):
        row_idx = i // 2
        col_idx = i % 2
        cell = qt.cell(row_idx, col_idx)
        # Pregunta con negrita parcial
        p = cell.paragraphs[0]
        p.add_run(f"{i+1}- ¿Qué es la ")
        bold_run = p.add_run("onda")
        bold_run.bold = True
        p.add_run(f" número {i+1}?")

        cell.add_paragraph(f"a) Opción A")
        cell.add_paragraph(f"b) Opción B")
        cell.add_paragraph(f"c) Opción C")
        cell.add_paragraph(f"d) Opción D")

    doc.save(path)


def test_build_docx_creates_file():
    with tempfile.TemporaryDirectory() as tmpdir:
        source = os.path.join(tmpdir, "original.docx")
        output = os.path.join(tmpdir, "forma_a.docx")
        _create_source_exam(source)

        exam = parse_docx(source)
        forms = generate_forms(exam, n=1)

        build_docx(
            source_path=source,
            questions=forms[0],
            output_path=output,
            form_suffix="A",
            questions_table_index=exam.questions_table_index,
            questions_per_row=exam.questions_per_row,
        )

        assert os.path.exists(output)
        assert os.path.getsize(output) > 0


def test_build_docx_preserves_header_tables():
    """Las tablas de header/instrucciones deben permanecer intactas."""
    with tempfile.TemporaryDirectory() as tmpdir:
        source = os.path.join(tmpdir, "original.docx")
        output = os.path.join(tmpdir, "forma_a.docx")
        _create_source_exam(source)

        exam = parse_docx(source)
        forms = generate_forms(exam, n=1)

        build_docx(
            source_path=source,
            questions=forms[0],
            output_path=output,
            form_suffix="A",
            questions_table_index=exam.questions_table_index,
            questions_per_row=exam.questions_per_row,
        )

        out_doc = Document(output)
        # Tabla 0 debe tener "NOMBRE"
        assert "NOMBRE" in out_doc.tables[0].cell(0, 0).text
        # Tabla 1 debe tener "INSTRUCCIONES"
        assert "INSTRUCCIONES" in out_doc.tables[1].cell(0, 0).text


def test_build_docx_updates_form_label():
    """El label 'Forma X' debe actualizarse al nuevo sufijo."""
    with tempfile.TemporaryDirectory() as tmpdir:
        source = os.path.join(tmpdir, "original.docx")
        output = os.path.join(tmpdir, "forma_b.docx")
        _create_source_exam(source)

        exam = parse_docx(source)
        forms = generate_forms(exam, n=1)

        build_docx(
            source_path=source,
            questions=forms[0],
            output_path=output,
            form_suffix="B",
            questions_table_index=exam.questions_table_index,
            questions_per_row=exam.questions_per_row,
        )

        out_doc = Document(output)
        full_text = " ".join(p.text for p in out_doc.paragraphs)
        assert "Forma B" in full_text
        assert "Forma 1" not in full_text


def test_build_docx_contains_all_questions():
    """El documento generado debe contener todas las preguntas."""
    with tempfile.TemporaryDirectory() as tmpdir:
        source = os.path.join(tmpdir, "original.docx")
        output = os.path.join(tmpdir, "forma_a.docx")
        _create_source_exam(source)

        exam = parse_docx(source)
        forms = generate_forms(exam, n=1)

        build_docx(
            source_path=source,
            questions=forms[0],
            output_path=output,
            form_suffix="A",
            questions_table_index=exam.questions_table_index,
            questions_per_row=exam.questions_per_row,
        )

        # Re-parsear para verificar
        out_exam = parse_docx(output)
        assert len(out_exam.questions) == 6
        for q in out_exam.questions:
            assert len(q.alternatives) == 4


def test_build_docx_preserves_bold_formatting():
    """El formato negrita del enunciado debe preservarse."""
    with tempfile.TemporaryDirectory() as tmpdir:
        source = os.path.join(tmpdir, "original.docx")
        output = os.path.join(tmpdir, "forma_a.docx")
        _create_source_exam(source)

        exam = parse_docx(source)
        # No randomizar — usar preguntas en orden original para verificar formato
        questions = [q.clone() for q in exam.questions]

        build_docx(
            source_path=source,
            questions=questions,
            output_path=output,
            form_suffix="A",
            questions_table_index=exam.questions_table_index,
            questions_per_row=exam.questions_per_row,
        )

        out_doc = Document(output)
        qt = out_doc.tables[exam.questions_table_index]
        first_cell = qt.cell(0, 0)
        first_para = first_cell.paragraphs[0]

        # Verificar que hay un run con bold
        has_bold = any(run.bold for run in first_para.runs)
        assert has_bold, "Se perdió el formato negrita en la primera celda"
```

- [ ] **Paso 3: Correr los tests**

Comando: `uv run pytest tests/test_doc_builder.py -v`
Esperado: PASS (5 tests)

- [ ] **Paso 4: Commit**

```bash
git add src/doc_builder.py tests/test_doc_builder.py
git commit -m "refactor: rewrite doc_builder to clone original and preserve styles"
```

---

### Tarea 5: Actualizar `main.py` para el nuevo flujo

**Archivos:**
- Modificar: `src/main.py`

- [ ] **Paso 1: Reescribir `src/main.py`**

```python
# src/main.py
"""Script principal del randomizador de pruebas.

Uso:
    python -m src.main <ruta_en_drive> <n_formas>

Ejemplo:
    python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 4
"""

import os
import shutil
import sys
import tempfile

from src.doc_builder import build_docx
from src.drive_client import download_docx, upload_docx
from src.parser import parse_docx
from src.randomizer import generate_forms


def get_form_suffix(index: int) -> str:
    """Genera el sufijo de la forma: A, B, C, ..., Z, AA, AB, ..."""
    if index < 26:
        return chr(ord("A") + index)
    else:
        first = chr(ord("A") + (index // 26) - 1)
        second = chr(ord("A") + (index % 26))
        return first + second


def main():
    """Punto de entrada principal del programa."""
    if len(sys.argv) < 3:
        print("Uso: python -m src.main <ruta_en_drive> <n_formas>")
        print('Ejemplo: python -m src.main "Guías/Ondas 2026. Ev. 1 F1.docx" 4')
        sys.exit(1)

    drive_path = sys.argv[1]
    n_forms = int(sys.argv[2])

    if n_forms < 1:
        print("❌ El número de formas debe ser al menos 1.")
        sys.exit(1)

    path_parts = drive_path.strip("/").split("/")
    filename = path_parts[-1]
    drive_folder = "/".join(path_parts[:-1]) if len(path_parts) > 1 else ""
    base_name = filename.rsplit(".docx", 1)[0]

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
            f"   Encontradas {len(exam.questions)} preguntas en tabla[{exam.questions_table_index}] "
            f"({exam.questions_per_row} columnas).\n"
        )

        if len(exam.questions) == 0:
            print("❌ No se detectaron preguntas en el documento.")
            print("   Verificá que las preguntas estén en una tabla y usen formato 'N-' o 'N.'")
            sys.exit(1)

        # 3. Generar N formas aleatorias
        print("🎲 Generando formas aleatorias...")
        forms = generate_forms(exam, n_forms)

        # 4. Construir y subir cada forma
        temp_dir = tempfile.mkdtemp()
        for i, form_questions in enumerate(forms):
            suffix = get_form_suffix(i)
            form_filename = f"{base_name} Forma {suffix}.docx"
            form_path = os.path.join(temp_dir, form_filename)

            print(f"\n📝 Generando: {form_filename}")
            build_docx(
                source_path=local_path,
                questions=form_questions,
                output_path=form_path,
                form_suffix=suffix,
                questions_table_index=exam.questions_table_index,
                questions_per_row=exam.questions_per_row,
            )

            print(f"⬆️  Subiendo a Drive...")
            upload_docx(form_path, drive_folder, form_filename)

            os.unlink(form_path)

        shutil.rmtree(temp_dir, ignore_errors=True)

    finally:
        if os.path.exists(local_path):
            os.unlink(local_path)

    print(f"\n✅ ¡Listo! Se generaron {n_forms} forma(s) exitosamente.")
    print("   Revisá tu carpeta en Google Drive.")


if __name__ == "__main__":
    main()
```

- [ ] **Paso 2: Verificar que el script corre sin argumentos**

Comando: `uv run python -m src.main`
Esperado: Mensaje de uso y exit code 1.

- [ ] **Paso 3: Commit**

```bash
git add src/main.py
git commit -m "refactor: update main.py for new clone-based build flow"
```

---

### Tarea 6: Test de integración end-to-end

**Archivos:**
- Modificar: `tests/test_integration.py`

- [ ] **Paso 1: Reescribir `tests/test_integration.py`**

```python
# tests/test_integration.py
"""Test de integración end-to-end — verifica el flujo completo sin Drive.

Crea un .docx con estructura de tabla realista, lo parsea, genera formas,
y verifica que los documentos generados preservan formato y contenido.
"""
import os
import tempfile
from docx import Document
from docx.shared import Pt, Inches
from src.parser import parse_docx
from src.randomizer import generate_forms
from src.doc_builder import build_docx


def _create_realistic_exam(path: str):
    """Crea un documento .docx con estructura realista de examen con tabla."""
    doc = Document()

    # Título con formato
    p_title = doc.add_paragraph()
    run_title = p_title.add_run("Prueba 1. Ondas y sus características")
    run_title.bold = True
    run_title.font.size = Pt(14)

    # Forma
    p_form = doc.add_paragraph()
    run_form = p_form.add_run("Forma 1")
    run_form.bold = True
    run_form.font.size = Pt(13)

    # Tabla datos alumno
    data_table = doc.add_table(rows=2, cols=4)
    data_table.cell(0, 0).paragraphs[0].add_run("NOMBRE").bold = True
    data_table.cell(0, 2).paragraphs[0].add_run("CURSO").bold = True
    data_table.cell(1, 0).paragraphs[0].add_run("FECHA").bold = True
    data_table.cell(1, 2).paragraphs[0].add_run("PUNTAJE").bold = True

    # Tabla instrucciones
    inst_table = doc.add_table(rows=1, cols=1)
    inst_cell = inst_table.cell(0, 0)
    inst_cell.paragraphs[0].add_run("INSTRUCCIONES:").bold = True
    inst_cell.add_paragraph("Lea con atención y seleccione su alternativa correcta.")

    # Párrafo de ítem
    p_item = doc.add_paragraph()
    r1 = p_item.add_run("Ítem I")
    r1.bold = True
    r1.italic = True
    r2 = p_item.add_run(". Marque la alternativa según corresponda (1 pto c/u)")
    r2.bold = True

    # Tabla de preguntas (5 filas x 2 cols = 10 preguntas)
    qt = doc.add_table(rows=5, cols=2)

    questions = [
        ("¿Cuál es la velocidad de la luz?", ["3x10^8 m/s", "3x10^6 m/s", "3x10^10 m/s", "3x10^4 m/s"]),
        ("¿Qué transmiten las ondas?", ["Materia", "Fuerza", "Energía", "Calor"]),
        ("¿Qué es el punto de equilibrio?", ["Posición de reposo", "Amplitud", "Valle", "Cresta"]),
        ("¿Qué representa la amplitud?", ["Distancia al equilibrio", "2 valles", "2 montes", "Tiempo"]),
        ("¿Qué es un monte?", ["Punto más alto", "Punto más bajo", "Equilibrio", "Nodo"]),
        ("¿Qué es un valle?", ["Punto más bajo", "Punto más alto", "Amplitud", "Fase"]),
        ("¿Qué tipo de onda es el sonido?", ["Longitudinal", "Transversal", "Electromagnética", "Estacionaria"]),
        ("¿Qué necesitan las ondas mecánicas?", ["Medio material", "Vacío", "Luz", "Electricidad"]),
        ("¿Qué es un pulso?", ["Perturbación aislada", "Onda periódica", "Vibración", "Resonancia"]),
        ("¿Cómo se propaga la luz en el vacío?", ["Como onda EM", "No se propaga", "Como sonido", "Como calor"]),
    ]

    for i, (q_text, alts) in enumerate(questions):
        row_idx = i // 2
        col_idx = i % 2
        cell = qt.cell(row_idx, col_idx)
        # Enunciado con alguna palabra en negrita
        p = cell.paragraphs[0]
        p.add_run(f"{i+1}- ")
        if i % 3 == 0:
            # Alguna pregunta con negrita parcial
            words = q_text.split(" ")
            mid = len(words) // 2
            p.add_run(" ".join(words[:mid]))
            bold_run = p.add_run(f" {words[mid]} ")
            bold_run.bold = True
            p.add_run(" ".join(words[mid+1:]))
        else:
            p.add_run(q_text)

        for j, alt_text in enumerate(alts):
            letter = chr(ord("a") + j)
            cell.add_paragraph(f"{letter}) {alt_text}")

    # Footer
    doc.add_paragraph("")
    p_footer = doc.add_paragraph()
    r_footer = p_footer.add_run("¡Has llegado al final de la prueba!")
    r_footer.bold = True
    r_footer.font.size = Pt(15)

    doc.save(path)


def test_full_pipeline():
    """Test completo: crear docx → parsear → randomizar → generar formas."""
    with tempfile.TemporaryDirectory() as tmpdir:
        source_path = os.path.join(tmpdir, "original.docx")
        _create_realistic_exam(source_path)

        # Parsear
        exam = parse_docx(source_path)
        assert len(exam.questions) == 10
        assert exam.questions_per_row == 2

        # Generar 3 formas
        forms = generate_forms(exam, n=3)
        assert len(forms) == 3

        # Construir archivos
        output_paths = []
        for i, form_questions in enumerate(forms):
            suffix = chr(ord("A") + i)
            out_path = os.path.join(tmpdir, f"original Forma {suffix}.docx")
            build_docx(
                source_path=source_path,
                questions=form_questions,
                output_path=out_path,
                form_suffix=suffix,
                questions_table_index=exam.questions_table_index,
                questions_per_row=exam.questions_per_row,
            )
            output_paths.append(out_path)

        # Verificar que los archivos existen
        for out_path in output_paths:
            assert os.path.exists(out_path)
            assert os.path.getsize(out_path) > 0

        # Re-parsear para verificar contenido
        for out_path in output_paths:
            out_exam = parse_docx(out_path)
            assert len(out_exam.questions) == 10
            for q in out_exam.questions:
                assert len(q.alternatives) == 4

        # Verificar que las formas tienen orden distinto
        form_orders = []
        for out_path in output_paths:
            parsed = parse_docx(out_path)
            order = tuple(q.text for q in parsed.questions)
            form_orders.append(order)
        unique_orders = set(form_orders)
        assert len(unique_orders) >= 2


def test_preserves_header_and_footer():
    """Verifica que el header, tablas de datos e instrucciones se preservan."""
    with tempfile.TemporaryDirectory() as tmpdir:
        source_path = os.path.join(tmpdir, "original.docx")
        _create_realistic_exam(source_path)

        exam = parse_docx(source_path)
        forms = generate_forms(exam, n=1)

        out_path = os.path.join(tmpdir, "forma_a.docx")
        build_docx(
            source_path=source_path,
            questions=forms[0],
            output_path=out_path,
            form_suffix="A",
            questions_table_index=exam.questions_table_index,
            questions_per_row=exam.questions_per_row,
        )

        out_doc = Document(out_path)

        # Header: título debe existir
        all_text = " ".join(p.text for p in out_doc.paragraphs)
        assert "Prueba 1" in all_text
        assert "Forma A" in all_text
        assert "final de la prueba" in all_text

        # Tabla de datos del alumno
        assert "NOMBRE" in out_doc.tables[0].cell(0, 0).text
        assert "CURSO" in out_doc.tables[0].cell(0, 2).text

        # Tabla de instrucciones
        assert "INSTRUCCIONES" in out_doc.tables[1].cell(0, 0).text


def test_preserves_formatting_in_output():
    """Verifica que negrita/cursiva se preservan en el output."""
    with tempfile.TemporaryDirectory() as tmpdir:
        source_path = os.path.join(tmpdir, "original.docx")
        _create_realistic_exam(source_path)

        exam = parse_docx(source_path)
        # Clonar sin randomizar para poder verificar formato
        questions = [q.clone() for q in exam.questions]

        out_path = os.path.join(tmpdir, "forma_a.docx")
        build_docx(
            source_path=source_path,
            questions=questions,
            output_path=out_path,
            form_suffix="A",
            questions_table_index=exam.questions_table_index,
            questions_per_row=exam.questions_per_row,
        )

        out_doc = Document(out_path)

        # El título debe tener bold
        title_para = out_doc.paragraphs[0]
        assert any(run.bold for run in title_para.runs)

        # Verificar que alguna pregunta n la tabla tiene bold parcial
        qt = out_doc.tables[exam.questions_table_index]
        first_cell = qt.cell(0, 0)
        first_para = first_cell.paragraphs[0]
        has_bold_run = any(run.bold for run in first_para.runs if run.text.strip())
        has_non_bold_run = any(not run.bold for run in first_para.runs if run.text.strip())
        # Debe tener mezcla de bold y no-bold (formato parcial)
        assert has_bold_run and has_non_bold_run
```

- [ ] **Paso 2: Correr todos los tests**

Comando: `uv run pytest tests/ -v`
Esperado: PASS (todos los tests — ~20 tests en total)

- [ ] **Paso 3: Commit**

```bash
git add tests/test_integration.py
git commit -m "test: rewrite integration tests for clone-based builder"
```

---

### Tarea 7: Eliminar `tests/test_models.py` obsoleto y limpiar

**Archivos:**
- Modificar: `tests/test_models.py`

- [ ] **Paso 1: Reescribir `tests/test_models.py` para nuevos modelos**

```python
# tests/test_models.py
"""Tests de los modelos v2 — XML-based."""
import os
import tempfile
from docx import Document
from src.parser import parse_docx
from src.models import Alternative, Question, Exam


def _create_simple_exam() -> Exam:
    """Helper: crea un exam simple parseando un docx temporal."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name

    doc = Document()
    qt = doc.add_table(rows=1, cols=1)
    cell = qt.cell(0, 0)
    p = cell.paragraphs[0]
    p.add_run("1- ")
    bold_run = p.add_run("Pregunta")
    bold_run.bold = True
    p.add_run(" de prueba")
    cell.add_paragraph("a) Opción A")
    cell.add_paragraph("b) Opción B")
    doc.save(path)

    exam = parse_docx(path)
    os.unlink(path)
    return exam


def test_question_text_property():
    """El .text debe devolver el texto plano del enunciado."""
    exam = _create_simple_exam()
    q = exam.questions[0]
    assert "Pregunta" in q.text
    assert "prueba" in q.text


def test_alternative_text_property():
    """El .text debe devolver el texto plano de la alternativa."""
    exam = _create_simple_exam()
    q = exam.questions[0]
    assert q.alternatives[0].text.strip()  # No vacío
    assert "Opción A" in q.alternatives[0].text or "a)" in q.alternatives[0].text


def test_question_clone_is_independent():
    """Clonar una pregunta no debe afectar el original."""
    exam = _create_simple_exam()
    q = exam.questions[0]
    original_text = q.text
    cloned = q.clone()

    # Modificar el clon no afecta al original
    cloned.number = 99
    assert q.number == 1
    assert q.text == original_text


def test_alternative_clone_is_independent():
    """Clonar una alternativa no debe afectar el original."""
    exam = _create_simple_exam()
    alt = exam.questions[0].alternatives[0]
    original_letter = alt.letter
    cloned = alt.clone()

    cloned.letter = "z"
    assert alt.letter == original_letter
```

- [ ] **Paso 2: Correr los tests de modelos**

Comando: `uv run pytest tests/test_models.py -v`
Esperado: PASS (4 tests)

- [ ] **Paso 3: Correr TODA la suite**

Comando: `uv run pytest tests/ -v`
Esperado: PASS (todos)

- [ ] **Paso 4: Commit**

```bash
git add tests/test_models.py
git commit -m "test: rewrite model tests for XML-based models"
```

---

### Tarea 8: Limpiar archivos de diagnóstico

**Archivos:**
- Eliminar: `analyze_doc.py`
- Eliminar: `list_drive.py`

- [ ] **Paso 1: Eliminar scripts de diagnóstico**

```bash
rm analyze_doc.py list_drive.py
```

- [ ] **Paso 2: Commit final**

```bash
git add -A
git commit -m "chore: remove diagnostic scripts, v2 complete"
```

---

## Resumen de tareas

| Tarea | Descripción | Archivos |
|-------|-------------|----------|
| 1 | Rediseñar modelos (XML-based) | `src/models.py` |
| 2 | Reescribir parser (extrae de tablas) | `src/parser.py`, `tests/test_parser.py` |
| 3 | Adaptar randomizador | `src/randomizer.py`, `tests/test_randomizer.py` |
| 4 | Reescribir doc_builder (clona original) | `src/doc_builder.py`, `tests/test_doc_builder.py` |
| 5 | Actualizar main.py | `src/main.py` |
| 6 | Test de integración e2e | `tests/test_integration.py` |
| 7 | Tests de modelos actualizados | `tests/test_models.py` |
| 8 | Limpiar scripts de diagnóstico | `analyze_doc.py`, `list_drive.py` |

## Decisiones de diseño clave

1. **Clonar en vez de recrear:** Se copia el .docx original byte a byte y solo se modifica la tabla de preguntas. Esto preserva TODO: logos, headers, footers, estilos de página, márgenes, fuentes, colores, tablas de datos.

2. **XML directo:** Los modelos almacenan elementos `lxml` (los `<w:p>` originales) en vez de texto plano. Permite manipular el contenido sin perder formato.

3. **Detección automática de tabla:** El parser busca la tabla con más filas (heurística robusta para el caso de uso).

4. **Formato de preguntas flexible:** El regex `^(\d+)\s*[-–—.]\s*` soporta tanto `N-` como `N.` y variantes con guión largo.