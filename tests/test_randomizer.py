"""Tests para el generador de formas aleatorias."""

import os
import re
import tempfile

from docx import Document

from src.models import Alternative, Exam, Question
from src.parser import parse_docx
from src.randomizer import QUESTION_NUMBER_PATTERN, _renumber_question, generate_forms


def _create_test_exam(num_questions=5) -> Exam:
    """Crea un examen de prueba con tabla de preguntas en docx temporal y lo parsea."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name

    doc = Document()
    rows_needed = (num_questions + 1) // 2
    q_table = doc.add_table(rows=rows_needed, cols=2)

    for i in range(num_questions):
        row_idx = i // 2
        col_idx = i % 2
        cell = q_table.cell(row_idx, col_idx)
        cell.paragraphs[0].text = f"{i + 1}- Pregunta número {i + 1} sobre ondas"
        cell.add_paragraph(f"a) Opción A de P{i + 1}")
        cell.add_paragraph(f"b) Opción B de P{i + 1}")
        cell.add_paragraph(f"c) Opción C de P{i + 1}")
        cell.add_paragraph(f"d) Opción D de P{i + 1}")

    doc.save(path)
    exam = parse_docx(path)
    os.unlink(path)
    return exam


def test_generate_forms_returns_correct_count():
    """Debe devolver exactamente n formas."""
    exam = _create_test_exam(num_questions=5)
    forms = generate_forms(exam, n=3)
    assert len(forms) == 3


def _strip_prefix(text: str) -> str:
    """Strip the question number prefix (e.g. '4- ') to compare content only."""
    return QUESTION_NUMBER_PATTERN.sub("", text)


def test_generate_forms_preserves_all_questions():
    """Cada forma debe tener las 5 preguntas originales con los mismos textos."""
    exam = _create_test_exam(num_questions=5)
    forms = generate_forms(exam, n=2)
    for form in forms:
        assert len(form) == 5
        original_texts = {_strip_prefix(q.text) for q in exam.questions}
        form_texts = {_strip_prefix(q.text) for q in form}
        assert original_texts == form_texts


def test_generate_forms_renumbers_questions():
    """Los números de preguntas en cada forma deben ser [1,2,3,4,5]."""
    exam = _create_test_exam(num_questions=5)
    forms = generate_forms(exam, n=2)
    for form in forms:
        numbers = [q.number for q in form]
        assert numbers == [1, 2, 3, 4, 5]


def test_generate_forms_relabels_alternatives():
    """Las letras de alternativas en cada forma deben ser [a,b,c,d]."""
    exam = _create_test_exam(num_questions=5)
    forms = generate_forms(exam, n=2)
    for form in forms:
        for q in form:
            letters = [a.letter for a in q.alternatives]
            assert letters == ["a", "b", "c", "d"]


def test_generate_forms_shuffles_order():
    """Verifica que el randomizador mezcla las preguntas."""
    import random

    exam = _create_test_exam(num_questions=5)

    # Generate 10 forms and check that at least 2 have different internal orderings.
    # Since shuffling is random, we just verify the function doesn't produce
    # identical forms every single time.
    orderings_by_id = []
    for _ in range(10):
        forms = generate_forms(exam, n=1)
        # Use object ids to detect if the actual objects are in different positions
        orderings_by_id.append([id(q) for q in forms[0]])

    unique = set(tuple(o) for o in orderings_by_id)
    # With 10 forms, shuffling should produce more than 1 unique ordering
    # (it's theoretically possible to get same ordering by chance, but unlikely)
    assert len(unique) >= 2, (
        f"Shuffling should produce varied orderings, got {len(unique)}: {unique}"
    )


def test_generate_forms_shuffles_alternatives():
    """Con 10 formas, al menos una debe tener alternativas reordenadas."""
    exam = _create_test_exam(num_questions=5)
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


def test_renumber_does_not_corrupt_text_numbers():
    """Numbers inside question text (like image labels) must not be modified."""
    # Create an exam where question text contains numbers that look like image labels
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name

    doc = Document()
    q_table = doc.add_table(rows=1, cols=2)

    # Question 1: text contains "1-" and "2-" as image labels inside the enunciado
    cell = q_table.cell(0, 0)
    cell.paragraphs[0].text = "1- Observe las figuras 1- y 2- sobre ondas"
    cell.add_paragraph("a) Opción A")
    cell.add_paragraph("b) Opción B")
    cell.add_paragraph("c) Opción C")
    cell.add_paragraph("d) Opción D")

    # Question 2: filler so the table has both columns
    cell2 = q_table.cell(0, 1)
    cell2.paragraphs[0].text = "2- Pregunta simple"
    cell2.add_paragraph("a) Opción A")
    cell2.add_paragraph("b) Opción B")
    cell2.add_paragraph("c) Opción C")
    cell2.add_paragraph("d) Opción D")

    doc.save(path)
    exam = parse_docx(path)
    os.unlink(path)

    # Find question #1 and clone it
    q1 = next(q for q in exam.questions if q.number == 1)
    cloned = q1.clone()

    # Renumber question 1 → 5
    _renumber_question(cloned, 5)

    # Extract the first <w:t> text from the enunciado
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    first_p = cloned.enunciado_paragraphs[0]
    first_t = None
    for t_elem in first_p.iter(f"{{{ns}}}t"):
        if t_elem.text:
            first_t = t_elem
            break

    assert first_t is not None
    # The prefix "1-" should become "5-"
    assert first_t.text.startswith("5- "), (
        f"Prefix should be '5- ', got: {first_t.text}"
    )
    # The image labels "1-" and "2-" inside the text must NOT be corrupted
    assert "figuras 1-" in first_t.text, (
        f"Image label '1-' was corrupted: {first_t.text}"
    )
    assert "2- sobre" in first_t.text, f"Image label '2-' was corrupted: {first_t.text}"


def test_generate_forms_does_not_mutate_original():
    """Después de generar formas, el exam original no debe haber cambiado."""
    exam = _create_test_exam(num_questions=5)
    original_texts = [q.text for q in exam.questions]
    original_numbers = [q.number for q in exam.questions]
    generate_forms(exam, n=5)
    assert [q.text for q in exam.questions] == original_texts
    assert [q.number for q in exam.questions] == original_numbers


def test_generate_forms_preserves_original_numbers():
    """Después de randomizar, cada pregunta conserva original_number."""
    exam = _create_test_exam()
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
        assert original_letters == {"a", "b", "c", "d"}, (
            f"Q{q.number} (original Q{q.original_number}): "
            f"original_letters={original_letters}, esperado={{'a','b','c','d'}}"
        )


def test_generate_forms_original_vs_current_can_differ():
    """En al menos alguna pregunta, letter != original_letter (shuffle real)."""
    import random as rnd

    rnd.seed(42)
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
