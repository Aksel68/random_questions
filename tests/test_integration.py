"""Test de integración que verifica el flujo completo sin Google Drive."""
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
        source_path = os.path.join(tmpdir, "examen_original.docx")
        _create_full_exam_docx(source_path)

        exam = parse_docx(source_path)
        assert len(exam.questions) == 5
        assert len(exam.header_paragraphs) >= 1

        forms = generate_forms(exam, n=3)
        assert len(forms) == 3

        output_paths = []
        for i, form in enumerate(forms):
            suffix = chr(ord("A") + i)
            out_path = os.path.join(tmpdir, f"examen_original Forma {suffix}.docx")
            build_docx(form, out_path)
            output_paths.append(out_path)

        for out_path in output_paths:
            assert os.path.exists(out_path)
            assert os.path.getsize(out_path) > 0

            re_parsed = parse_docx(out_path)
            assert len(re_parsed.questions) == 5
            for q in re_parsed.questions:
                assert len(q.alternatives) == 5

        form_orders = []
        for out_path in output_paths:
            parsed = parse_docx(out_path)
            order = tuple(q.text for q in parsed.questions)
            form_orders.append(order)

        unique_orders = set(form_orders)
        assert len(unique_orders) >= 2


def test_roundtrip_preserves_content():
    """Verifica que el contenido se preserva tras parse → build → parse."""
    with tempfile.TemporaryDirectory() as tmpdir:
        source_path = os.path.join(tmpdir, "original.docx")
        _create_full_exam_docx(source_path)

        exam = parse_docx(source_path)
        original_texts = {q.text for q in exam.questions}
        original_alt_counts = {len(q.alternatives) for q in exam.questions}

        rebuilt_path = os.path.join(tmpdir, "rebuilt.docx")
        build_docx(exam, rebuilt_path)

        rebuilt_exam = parse_docx(rebuilt_path)
        rebuilt_texts = {q.text for q in rebuilt_exam.questions}

        assert original_texts == rebuilt_texts
        assert all(len(q.alternatives) == 5 for q in rebuilt_exam.questions)
