"""Generador de formas aleatorias a partir de un examen.

Produce N versiones del examen con preguntas y alternativas en orden aleatorio.
Trabaja con elementos XML clonados — preserva todo el formato original.
"""

import random
import re
from copy import deepcopy

from src.models import Alternative, Exam, Question

# Patrones para reemplazar número de pregunta y letra de alternativa en el XML
QUESTION_NUMBER_PATTERN = re.compile(r"^(\d+)\s*[-–—.]\s*")
ALTERNATIVE_LETTER_PATTERN = re.compile(r"^([a-zA-Z])\)\s*")


def _get_first_text_element(p_element):
    """Obtiene el primer elemento <w:t> con texto de un párrafo.

    Args:
        p_element: Elemento XML <w:p>.

    Returns:
        El primer <w:t> que contiene texto, o None.
    """
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    for t_elem in p_element.iter(f"{{{ns}}}t"):
        if t_elem.text:
            return t_elem
    return None


def _renumber_question(question: Question, new_number: int) -> Question:
    """Cambia el número de una pregunta en su XML.

    Actualiza únicamente el patrón "N-" o "N." en el primer <w:t> que
    coincida dentro de los párrafos del enunciado. No modifica ningún otro
    elemento de texto para evitar corromper números que aparezcan en el
    contenido de la pregunta (por ejemplo etiquetas de imagen como "1-").

    Args:
        question: La pregunta (ya clonada).
        new_number: El nuevo número a asignar.

    Returns:
        La misma pregunta con número actualizado.
    """
    old_number = question.number
    question.number = new_number

    if not question.enunciado_paragraphs:
        return question

    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # Only replace the question number prefix "N-" in the first matching <w:t>
    for p_elem in question.enunciado_paragraphs:
        for t_elem in p_elem.iter(f"{{{ns}}}t"):
            if t_elem.text and QUESTION_NUMBER_PATTERN.match(t_elem.text):
                t_elem.text = QUESTION_NUMBER_PATTERN.sub(
                    f"{new_number}- ", t_elem.text
                )
                return question

    return question


def _relabel_alternative(alt: Alternative, new_letter: str) -> Alternative:
    """Cambia la letra de una alternativa en su XML.

    Busca el patrón "x)" en el primer <w:t> del primer párrafo
    y lo reemplaza con la nueva letra.

    Args:
        alt: La alternativa (ya clonada).
        new_letter: La nueva letra (a, b, c, d...).

    Returns:
        La misma alternativa con letra actualizada.
    """
    alt.letter = new_letter

    if not alt.paragraphs:
        return alt

    first_p = alt.paragraphs[0]
    t_elem = _get_first_text_element(first_p)
    if t_elem is not None and t_elem.text:
        t_elem.text = ALTERNATIVE_LETTER_PATTERN.sub(f"{new_letter}) ", t_elem.text)

    return alt


def _shuffle_alternatives(alternatives: list[Alternative]) -> list[Alternative]:
    """Reordena las alternativas y reasigna letras.

    Args:
        alternatives: Lista de alternativas (ya clonadas).

    Returns:
        Lista de alternativas en orden aleatorio con letras reasignadas.
    """
    shuffled = list(alternatives)
    random.shuffle(shuffled)
    for i, alt in enumerate(shuffled):
        new_letter = chr(ord("a") + i)
        _relabel_alternative(alt, new_letter)
    return shuffled


def generate_forms(exam: Exam, n: int) -> list[list[Question]]:
    """Genera N listas de preguntas randomizadas.

    Cada forma es una lista de Questions con:
    - Orden de preguntas aleatorio
    - Números reasignados (1, 2, 3...)
    - Alternativas en orden aleatorio con letras reasignadas

    Args:
        exam: El examen original parseado.
        n: Número de formas a generar.

    Returns:
        Lista de N listas de preguntas randomizadas (con XML clonado).

    Raises:
        ValueError: Si n < 1.
    """
    if n < 1:
        raise ValueError("El número de formas debe ser al menos 1")

    forms: list[list[Question]] = []

    for _ in range(n):
        # Clonar todas las preguntas
        cloned_questions = [q.clone() for q in exam.questions]

        # Mezclar orden de preguntas
        random.shuffle(cloned_questions)

        # Renumerar y mezclar alternativas
        for i, q in enumerate(cloned_questions):
            _renumber_question(q, i + 1)
            q.alternatives = _shuffle_alternatives(q.alternatives)

        forms.append(cloned_questions)

    return forms
