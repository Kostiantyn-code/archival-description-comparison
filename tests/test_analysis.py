from __future__ import annotations

import unittest
from pathlib import Path

from src.analysis import (
    category_rows,
    classification_coverage_rows,
    topic_unclassified_rows,
)
from src.models import Category, Dataset, Record


def make_dataset(identifier: str) -> Dataset:
    return Dataset(
        identifier,
        f"Масив {identifier}",
        identifier.upper(),
        f"Архів, ф. {identifier}",
        Path(f"{identifier}.xlsx"),
    )


def make_record(dataset: Dataset, number: int) -> Record:
    return Record(
        dataset_id=dataset.id,
        dataset_label=dataset.label,
        file_name=dataset.path.name,
        sheet_name="Опис 1",
        excel_row=number + 1,
        archive="Архів",
        fond=dataset.id,
        inventory="1",
        case_id=str(number),
        title=f"Заголовок {number}",
        dates_raw="1850",
        pages_raw="1",
        notes="",
    )


class AnalysisTests(unittest.TestCase):
    def setUp(self) -> None:
        self.left = make_dataset("left")
        self.right = make_dataset("right")
        self.category = Category(
            id="culture",
            label="Культура",
            macroblock="cultural",
            minimum_score=1,
            strong_phrases=[],
            strong_terms=[],
            context_only_phrases=[],
            context_only_terms=[],
            contextual_terms=[],
            context_rules=[],
            exclude_phrases=[],
            review_terms=[],
        )

    def test_category_percent_uses_each_dataset_total(self):
        records = [
            *(make_record(self.left, number) for number in range(1, 5)),
            *(make_record(self.right, number) for number in range(1, 3)),
        ]
        records[0].categories = ["culture"]
        records[1].categories = ["culture"]
        records[4].categories = ["culture"]

        rows = category_rows(
            [self.left, self.right], records, [self.category]
        )

        self.assertEqual(rows[0]["dataset_titles_total"], 4)
        self.assertEqual(rows[0]["cases"], 2)
        self.assertEqual(rows[0]["percent_of_titles"], 50.0)
        self.assertEqual(rows[1]["dataset_titles_total"], 2)
        self.assertEqual(rows[1]["cases"], 1)
        self.assertEqual(rows[1]["percent_of_titles"], 50.0)

    def test_classification_states_are_exclusive_and_sum_to_total(self):
        records = [make_record(self.left, number) for number in range(1, 5)]
        records[0].categories = ["culture"]
        records[1].categories = ["culture"]
        records[2].context_categories = ["education"]

        row = classification_coverage_rows([self.left], records)[0]

        self.assertEqual(row["analyzable_titles"], 4)
        self.assertEqual(row["subject_classified"], 2)
        self.assertEqual(row["context_only"], 1)
        self.assertEqual(row["unclassified"], 1)
        self.assertEqual(row["subject_classified_percent"], 50.0)
        self.assertEqual(row["context_only_percent"], 25.0)
        self.assertEqual(row["unclassified_percent"], 25.0)
        self.assertAlmostEqual(
            row["subject_classified_percent"]
            + row["context_only_percent"]
            + row["unclassified_percent"],
            100.0,
        )

        detail = topic_unclassified_rows(records)
        self.assertEqual(len(detail), 2)
        self.assertEqual(
            [item["classification_status"] for item in detail],
            ["Лише контекст", "Не класифіковано"],
        )


if __name__ == "__main__":
    unittest.main()
