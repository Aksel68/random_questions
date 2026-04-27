import copy

from src.randomizer import generate_forms

from src.models import Alternative, Exam, Question


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
        original_texts = {q.text for q in exam.questions}
        form_texts = {q.text for q in form.questions}
        assert original_texts == form_texts


def test_generate_forms_renumbers_questions():
    exam = _make_exam()
    forms = generate_forms(exam, n=2)
    for form in forms:
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
    exam = _make_exam()
    forms = generate_forms(exam, n=10)
    original_order = [q.text for q in exam.questions]
    any_different = any(
        [q.text for q in form.questions] != original_order for form in forms
    )
    assert any_different


def test_generate_forms_shuffles_alternatives():
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
    assert any_different_alts


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
