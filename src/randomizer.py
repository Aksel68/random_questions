"""Generador de formas aleatorias a partir de un examen."""

import copy
import random

from src.models import Alternative, Exam, Question


def _shuffle_alternatives(alternatives: list[Alternative]) -> list[Alternative]:
    """Reordena las alternativas y reasigna las letras a, b, c, d..."""
    shuffled = copy.deepcopy(alternatives)
    random.shuffle(shuffled)
    for i, alt in enumerate(shuffled):
        alt.letter = chr(ord("a") + i)
    return shuffled


def _shuffle_questions(questions: list[Question]) -> list[Question]:
    """Reordena las preguntas, reasigna números, y shufflea sus alternativas."""
    shuffled = copy.deepcopy(questions)
    random.shuffle(shuffled)
    for i, q in enumerate(shuffled):
        q.number = i + 1
        q.alternatives = _shuffle_alternatives(q.alternatives)
    return shuffled


def generate_forms(exam: Exam, n: int) -> list[Exam]:
    """Genera N formas aleatorias del examen."""
    if n < 1:
        raise ValueError("El número de formas debe ser al menos 1")

    forms: list[Exam] = []
    for i in range(n):
        shuffled_questions = _shuffle_questions(exam.questions)
        form = Exam(
            title=exam.title,
            questions=shuffled_questions,
            header_paragraphs=list(exam.header_paragraphs),
        )
        forms.append(form)

    return forms
