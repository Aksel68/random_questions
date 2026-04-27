"""Generador de archivos .docx a partir de un examen."""

import io

from docx import Document
from docx.shared import Inches

from src.models import Exam, Question


def _add_image_to_paragraph(paragraph, image_data: tuple[bytes, str]):
    """Agrega una imagen a un párrafo."""
    image_bytes, content_type = image_data
    image_stream = io.BytesIO(image_bytes)
    run = paragraph.add_run()
    run.add_picture(image_stream, width=Inches(4.0))


def build_docx(exam: Exam, output_path: str) -> None:
    """Construye un archivo .docx a partir de un examen."""
    doc = Document()

    # Header
    for header_text in exam.header_paragraphs:
        doc.add_paragraph(header_text)

    if exam.header_paragraphs:
        doc.add_paragraph("")

    # Preguntas
    for i, question in enumerate(exam.questions):
        text_lines = question.text.split("\n")
        q_paragraph = doc.add_paragraph(f"{question.number}- {text_lines[0]}")

        for extra_line in text_lines[1:]:
            doc.add_paragraph(extra_line)

        for image_data in question.images:
            img_paragraph = doc.add_paragraph()
            _add_image_to_paragraph(img_paragraph, image_data)

        for alt in question.alternatives:
            alt_text = f"{alt.letter}) {alt.text}"
            alt_paragraph = doc.add_paragraph(alt_text)

            if alt.image:
                _add_image_to_paragraph(alt_paragraph, alt.image)

        if i < len(exam.questions) - 1:
            doc.add_paragraph("")

    doc.save(output_path)
