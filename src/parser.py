"""Parser de archivos .docx con formato de pruebas de alternativas."""
import re
from docx import Document
from src.models import Alternative, Question, Exam

QUESTION_PATTERN = re.compile(r"^(\d+)\s*[-–—]\s*(.*)$")
ALTERNATIVE_PATTERN = re.compile(r"^([a-zA-Z])\)\s*(.*)$")


def _extract_images_from_paragraph(paragraph) -> list[tuple[bytes, str]]:
    """Extrae imágenes embebidas en un párrafo."""
    images = []
    for run in paragraph.runs:
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
    """Parsea un archivo .docx de prueba de alternativas."""
    doc = Document(file_path)
    questions: list[Question] = []
    header_paragraphs: list[str] = []
    current_question: Question | None = None
    found_first_question = False

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        images = _extract_images_from_paragraph(paragraph)

        q_match = QUESTION_PATTERN.match(text)
        if q_match:
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

        a_match = ALTERNATIVE_PATTERN.match(text)
        if a_match and current_question is not None:
            letter = a_match.group(1).lower()
            alt_text = a_match.group(2).strip()
            alt_image = images[0] if images else None
            current_question.alternatives.append(
                Alternative(letter=letter, text=alt_text, image=alt_image)
            )
            continue

        if not found_first_question:
            if text:
                header_paragraphs.append(text)
        elif current_question is not None and text:
            current_question.text += "\n" + text
            current_question.images.extend(images)

    if current_question is not None:
        questions.append(current_question)

    title = header_paragraphs[0] if header_paragraphs else "Examen"

    return Exam(
        title=title,
        questions=questions,
        header_paragraphs=header_paragraphs,
    )
