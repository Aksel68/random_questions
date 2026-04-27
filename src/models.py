# src/models.py
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Alternative:
    """Una alternativa de una pregunta de selección múltiple."""

    letter: str
    text: str
    image: Optional[tuple[bytes, str]] = None


@dataclass
class Question:
    """Una pregunta completa con enunciado y alternativas."""

    number: int
    text: str
    alternatives: list[Alternative] = field(default_factory=list)
    images: list[tuple[bytes, str]] = field(default_factory=list)


@dataclass
class Exam:
    """Un examen completo con título y preguntas."""

    title: str
    questions: list[Question] = field(default_factory=list)
    header_paragraphs: list[str] = field(default_factory=list)
