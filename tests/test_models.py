# tests/test_models.py
from src.models import Alternative, Exam, Question


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
