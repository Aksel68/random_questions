import os
import tempfile

from docx import Document

from src.models import Exam
from src.parser import parse_docx


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

        q1 = exam.questions[0]
        assert q1.number == 1
        assert "velocidad de la luz" in q1.text.lower()
        assert len(q1.alternatives) == 4
        assert q1.alternatives[0].letter == "a"
        assert "3 x 10^8" in q1.alternatives[0].text

        q2 = exam.questions[1]
        assert q2.number == 2
        assert "sonido" in q2.text.lower()
        assert len(q2.alternatives) == 4
    finally:
        os.unlink(path)


def test_parse_docx_extracts_header():
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        path = f.name
    try:
        create_sample_docx(path)
        exam = parse_docx(path)

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
