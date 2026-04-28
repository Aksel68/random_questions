# src/models.py
"""Modelos de datos para el randomizador de pruebas."""

from copy import deepcopy
from dataclasses import dataclass, field
from typing import Optional

from lxml import etree


@dataclass
class Alternative:
    letter: str
    paragraphs: list  # list of lxml elements (<w:p>)
    original_letter: str = ""

    def __post_init__(self):
        if not self.original_letter:
            self.original_letter = self.letter

    def clone(self) -> "Alternative":
        return Alternative(
            letter=self.letter,
            paragraphs=[deepcopy(p) for p in self.paragraphs],
            original_letter=self.original_letter,
        )

    @property
    def text(self) -> str:
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        parts = []
        for p in self.paragraphs:
            runs = p.findall(".//w:r/w:t", nsmap)
            parts.append("".join(r.text or "" for r in runs))
        return " ".join(parts).strip()


@dataclass
class Question:
    number: int
    enunciado_paragraphs: list  # list of lxml elements (<w:p>)
    alternatives: list[Alternative] = field(default_factory=list)
    trailing_empty_paragraphs: list = field(default_factory=list)
    original_number: int = 0

    def __post_init__(self):
        if self.original_number == 0:
            self.original_number = self.number

    def clone(self) -> "Question":
        return Question(
            number=self.number,
            enunciado_paragraphs=[deepcopy(p) for p in self.enunciado_paragraphs],
            alternatives=[alt.clone() for alt in self.alternatives],
            trailing_empty_paragraphs=[
                deepcopy(p) for p in self.trailing_empty_paragraphs
            ],
            original_number=self.original_number,
        )

    @property
    def text(self) -> str:
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        parts = []
        for p in self.enunciado_paragraphs:
            runs = p.findall(".//w:r/w:t", nsmap)
            parts.append("".join(r.text or "" for r in runs))
        return " ".join(parts).strip()


@dataclass
class Exam:
    source_path: str
    questions: list[Question] = field(default_factory=list)
    questions_table_index: int = 2
    questions_per_row: int = 2
