from __future__ import annotations

import unittest
from pathlib import Path

from src.analysis import (
    categories_by_description_rows,
    category_rows,
    chronology_by_description_rows,
    chronology_rows,
    classification_by_description_rows,
    classification_coverage_rows,
    description_rows,
    overview_rows,
    topic_unclassified_rows,
)
from src.models import Category, Dataset, Description, Record


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
    def test_scope_tables_keep_identical_sheet_names_separate_and_empty_scopes(self):
        left, right, empty = [make_dataset(name) for name in ("left", "right", "empty")]
        datasets = [left, right, empty]
        descriptions = [Description(d.id, d.path.name, "Опис 1", "Архів", d.id,
                                    "1", "uk", 4, d.reference)
                        for d in (right, empty, left)]
        records = [make_record(left, 1), make_record(right, 1),
                   make_record(left, 2), make_record(left, 3), make_record(left, 4)]
        records[0].categories = ["culture"]
        records[0].context_categories = ["education"]
        records[2].context_categories = ["education"]
        records[3].title = " "  # A case without an analyzable title.
        records[4].status = "withdrawn"
        before = [(r.dataset_id, r.sheet_name, r.title, list(r.categories)) for r in records]
        category = Category("culture", "Культура", "cultural", 1,
                            [], [], [], [], [], [], [], [])

        overview = overview_rows(datasets, descriptions, records)
        self.assertEqual([(r["cases"], r["analyzable_titles"], r["withdrawn"])
                          for r in overview], [(3, 2, 1), (1, 1, 0), (0, 0, 0)])
        coverage = classification_coverage_rows(datasets, records)
        self.assertEqual([(r["subject_classified"], r["context_only"], r["unclassified"])
                          for r in coverage], [(1, 1, 0), (0, 0, 1), (0, 0, 0)])
        detail = description_rows(descriptions, records, {d.id: d for d in datasets})
        self.assertEqual([r["dataset_id"] for r in detail], ["right", "empty", "left"])
        self.assertEqual([r["analyzable_titles"] for r in detail], [1, 0, 2])
        categories = categories_by_description_rows(datasets, descriptions, records, [category])
        self.assertEqual([r["cases"] for r in categories], [0, 0, 1])
        self.assertEqual([r["percent_of_dataset_titles"] for r in categories], [0, 0, 50])
        self.assertEqual([r["cases"] for r in category_rows(datasets, records, [category])], [1, 0, 0])
        self.assertEqual([(r.dataset_id, r.sheet_name, r.title, list(r.categories))
                          for r in records], before)

        # A new call must reflect edits rather than reuse stale grouped state.
        records[0].categories = []
        updated = classification_by_description_rows(datasets, descriptions, records)
        self.assertEqual(updated[-1]["context_only"], 2)
        self.assertEqual(updated[-1]["subject_classified"], 0)

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

    def test_description_breakdowns_reconcile_with_workbook_totals(self):
        def description(dataset: Dataset, sheet: str) -> Description:
            return Description(dataset.id, dataset.path.name, sheet, "Архів",
                               dataset.id, sheet, "uk", 5, f"Архів, оп. {sheet}")

        descriptions = [description(self.left, "Опис 1"),
                        description(self.left, "Опис 2"),
                        description(self.right, "Опис 1")]
        first = make_record(self.left, 1)
        first.start_year, first.end_year, first.categories = 1850, 1851, ["culture"]
        second = make_record(self.left, 2)
        second.sheet_name, second.start_year = "Опис 2", 1851
        second.context_categories = ["education"]
        third = make_record(self.left, 3)
        third.sheet_name, third.start_year = "Опис 2", 1852
        fourth = make_record(self.right, 1)
        fourth.start_year, fourth.categories = 1850, ["culture"]
        withdrawn = make_record(self.left, 4)
        withdrawn.sheet_name, withdrawn.status, withdrawn.start_year = "Опис 2", "withdrawn", 1850
        records = [first, second, third, fourth, withdrawn]

        yearly = chronology_by_description_rows(
            [self.left, self.right], descriptions, records, {"maximum_span_years": 100}
        )
        aggregate = chronology_rows([self.left, self.right], records, {"maximum_span_years": 100})
        for row in aggregate:
            for dataset in (self.left, self.right):
                self.assertEqual(
                    sum(part["cases"] for part in yearly
                        if part["year"] == row["year"] and part["dataset_id"] == dataset.id),
                    row[dataset.id],
                )
        self.assertEqual(len(yearly), 9)  # Three descriptions on one shared, contiguous axis.

        coverage = classification_by_description_rows([self.left, self.right], descriptions, records)
        totals = classification_coverage_rows([self.left, self.right], records)
        for total in totals:
            pieces = [part for part in coverage if part["dataset_id"] == total["dataset_id"]]
            for key in ("analyzable_titles", "subject_classified", "context_only", "unclassified"):
                self.assertEqual(sum(part[key] for part in pieces), total[key])
        self.assertEqual(coverage[1]["context_only_percent"], 50.0)

        parts = categories_by_description_rows(
            [self.left, self.right], descriptions, records, [self.category]
        )
        categories = category_rows([self.left, self.right], records, [self.category])
        self.assertEqual(
            sum(part["cases"] for part in parts if part["dataset_id"] == self.left.id),
            categories[0]["cases"],
        )
        self.assertEqual(parts[0]["percent_of_description_titles"], 100.0)
        self.assertEqual(parts[0]["percent_of_dataset_titles"], 33.33)


if __name__ == "__main__":
    unittest.main()
