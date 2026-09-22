"""Структури даних для порівняльного аналізу архівних описів."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


CATEGORY_RULE_FIELDS = (
    "strong_phrases",
    "strong_terms",
    "context_only_phrases",
    "context_only_terms",
    "contextual_terms",
    "context_rules",
    "exclude_phrases",
    "review_terms",
)


@dataclass
class Issue:
    level: str
    file_name: str
    sheet_name: str
    row: int | None
    field: str
    value: str
    message: str


@dataclass
class Category:
    id: str
    label: str
    macroblock: str
    minimum_score: int
    strong_phrases: list[str]
    strong_terms: list[str]
    context_only_phrases: list[str]
    context_only_terms: list[str]
    contextual_terms: list[str]
    context_rules: list[dict[str, Any]]
    exclude_phrases: list[str]
    review_terms: list[str]


@dataclass
class Dataset:
    id: str
    label: str
    short_label: str
    reference: str
    path: Path
    configured: bool = False
    minimum_year: int | None = None
    maximum_year: int | None = None


@dataclass
class Description:
    dataset_id: str
    file_name: str
    sheet_name: str
    archive: str
    fond: str
    inventory: str
    language: str
    header_row: int
    reference: str


@dataclass
class Record:
    dataset_id: str
    dataset_label: str
    file_name: str
    sheet_name: str
    excel_row: int
    archive: str
    fond: str
    inventory: str
    case_id: str
    title: str
    dates_raw: str
    pages_raw: str
    notes: str
    status: str = "case"
    section_label: str = ""
    section_year: int | None = None
    start_year: int | None = None
    end_year: int | None = None
    pages: int | None = None
    language: str = "uk"
    categories: list[str] = field(default_factory=list)
    category_labels: list[str] = field(default_factory=list)
    macroblocks: list[str] = field(default_factory=list)
    scores: dict[str, int] = field(default_factory=dict)
    evidence: dict[str, list[str]] = field(default_factory=dict)
    context_categories: list[str] = field(default_factory=list)
    context_category_labels: list[str] = field(default_factory=list)
    context_evidence: dict[str, list[str]] = field(default_factory=dict)
    review_flags: list[str] = field(default_factory=list)

    @property
    def uid(self) -> str:
        return f"{self.dataset_id}:{self.sheet_name}:{self.case_id or 'row-' + str(self.excel_row)}"

    @property
    def description_reference(self) -> str:
        parts = [self.archive]
        if self.fond:
            parts.append(f"ф. {self.fond}")
        if self.inventory:
            parts.append(f"оп. {self.inventory}")
        return ", ".join(part for part in parts if part)

    @property
    def analyzable(self) -> bool:
        return self.status == "case" and bool(self.title.strip())
