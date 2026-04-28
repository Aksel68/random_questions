"""Tests para src/main.py — extract_form_number y lógica de nomenclatura."""

import pytest

from src.main import extract_form_number


class TestExtractFormNumber:
    def test_simple_f1(self):
        prefix, num = extract_form_number("Ondas 2026. Ev. 1 F1.docx")
        assert prefix == "Ondas 2026. Ev. 1 F"
        assert num == 1

    def test_f5(self):
        prefix, num = extract_form_number("Prueba F5.docx")
        assert prefix == "Prueba F"
        assert num == 5

    def test_multi_digit(self):
        prefix, num = extract_form_number("Test F12.docx")
        assert prefix == "Test F"
        assert num == 12

    def test_no_match_raises(self):
        with pytest.raises(ValueError, match="No se encontró"):
            extract_form_number("Prueba sin forma.docx")

    def test_no_extension_raises(self):
        with pytest.raises(ValueError, match="No se encontró"):
            extract_form_number("Otro archivo.txt")

    def test_f10(self):
        prefix, num = extract_form_number("Guía F10.docx")
        assert prefix == "Guía F"
        assert num == 10

    def test_with_folder_path(self):
        # El path completo se procesa — el prefijo incluye las carpetas
        prefix, num = extract_form_number("Carpeta/Subcarpeta/Ondas F3.docx")
        assert prefix == "Carpeta/Subcarpeta/Ondas F"
        assert num == 3


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
