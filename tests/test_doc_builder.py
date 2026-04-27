import os
import tempfile

from docx import Document
from src.doc_builder import build_docx

from src.models import Alternative, Exam, Question


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

        assert "1-" in full_text
        assert "2-" in full_text
        assert "3-" in full_text
        assert "a)" in full_text
        assert "b)" in full_text
        assert "c)" in full_text
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

        question_lines = [t for t in texts if t and t[0].isdigit() and "-" in t[:4]]
        assert len(question_lines) == 3
    finally:
        os.unlink(output_path)
