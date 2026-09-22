from __future__ import annotations

import unittest
from pathlib import Path

from src.analysis import similarity_rows
from src.models import Dataset, Record


def record(dataset: Dataset, case_id: str, title: str) -> Record:
    return Record(
        dataset_id=dataset.id,
        dataset_label=dataset.label,
        file_name=dataset.path.name,
        sheet_name="Опис 1",
        excel_row=5,
        archive="Архів",
        fond="1",
        inventory="1",
        case_id=case_id,
        title=title,
        dates_raw="1918",
        pages_raw="1",
        notes="",
        start_year=1918,
        end_year=1918,
        pages=1,
    )


class SimilarityTests(unittest.TestCase):
    def test_related_titles_are_found_between_datasets(self):
        left = Dataset("left", "Лівий", "Лівий", "", Path("left.xlsx"))
        right = Dataset("right", "Правий", "Правий", "", Path("right.xlsx"))
        rows = similarity_rows(
            [left, right],
            [
                record(left, "1", "Листування про відкриття жіночого училища"),
                record(left, "2", "Справа про ремонт міського мосту"),
                record(right, "7", "Переписка об открытии женского училища"),
                record(right, "8", "Відомості про постачання зерна"),
            ],
            {
                "enabled": True,
                "minimum_score": 0.15,
                "matches_per_case": 1,
                "max_features": 5000,
                "character_ngram_min": 3,
                "character_ngram_max": 5,
            },
        )
        self.assertTrue(rows)
        self.assertIn("училища", rows[0]["title_a"] + rows[0]["title_b"])


if __name__ == "__main__":
    unittest.main()
