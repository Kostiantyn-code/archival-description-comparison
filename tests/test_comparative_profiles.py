from __future__ import annotations

import csv
import json
import tempfile
import unittest
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

from src.analysis import build_analysis
from src.comparative_profiles import REPORT_COLUMNS, build_comparative_profiles, write_theme_network
from src.document_types import match_document_types
from src.models import Category, Dataset, Description, Record
from src.reports import write_csv_reports, write_html_report, write_workbook


BASE = Path(__file__).resolve().parents[1]


class ComparativeProfileTests(unittest.TestCase):
    def setUp(self):
        self.datasets = [Dataset(key, key, key, "", Path(key + ".xlsx")) for key in ("a", "b", "empty")]
        self.descriptions = [
            Description("a", "a.xlsx", "Опис 1", "ДАМО", "230", "1", "uk", 5, "ДАМО, ф. 230, оп. 1"),
            Description("a", "a.xlsx", "ЦДІАК", "ЦДІАК", "356", "1", "ru", 5, "ЦДІАК, ф. 356, оп. 1"),
            Description("b", "b.xlsx", "Опис 1", "ЦДАВО", "1793", "1", "uk", 5, "ЦДАВО, ф. 1793, оп. 1"),
            Description("empty", "empty.xlsx", "Опис 1", "Архів", "1", "1", "uk", 5, ""),
        ]
        self.categories = [Category(key, key, block, 1, [], [], [], [], [], [], [], [])
                           for key, block in (("a", "one"), ("b", "one"), ("c", "two"), ("absent", "three"))]
        self.records = []
        def record(description, row, title, categories=(), status="case"):
            d = self.descriptions[description]
            item = Record(d.dataset_id, d.dataset_id, d.file_name, d.sheet_name, row,
                          d.archive, d.fond, d.inventory, "1", title, "1850", "10", "",
                          status=status, categories=list(categories))
            self.records.append(item)
            return item
        record(0, 6, "Листування, листи та рапорти", ["a", "a", "b", "c"])
        record(0, 7, "Заголовок без типу", ["a"])
        record(1, 6, "Переписка и рапорты", ["b"])
        record(1, 7, "Інший заголовок").context_categories = ["a", "b"]
        record(2, 6, "Рапорти", ["a", "b"])
        record(2, 7, "Невідомий тип", ["a", "c"])
        record(0, 8, "Рапорти", ["a", "b"], "withdrawn")
        record(0, 9, "   ", ["a", "b"])
        record(3, 6, "Рапорти", ["a", "b"], "withdrawn")
        self.result = build_comparative_profiles(self.datasets, self.descriptions, self.records, self.categories)

    def test_mentions_are_binary_per_title_and_use_all_analyzable_titles(self):
        row = next(r for r in self.result["document_types"]
                   if r["dataset_id"] == "a" and r["document_type_id"] == "correspondence")
        self.assertEqual((row["titles_total"], row["titles_with_type"], row["percent_of_titles"], row["per_1000_titles"]),
                         (4, 2, 50, 500))
        mentions = [r for r in self.result["document_type_mentions"] if r["document_type_id"] == "correspondence"]
        self.assertEqual(len(mentions), 2)
        self.assertEqual({r["archive"] for r in mentions}, {"ДАМО", "ЦДІАК"})
        for title in ("Справа про відкриття школи", "Дело о школе", "Листопад", "Листовий метал"):
            self.assertEqual(match_document_types(title), [])

    def test_pairs_lift_and_macroblocks_exclude_context_and_deduplicate(self):
        row = next(r for r in self.result["theme_links"] if r["dataset_id"] == "a"
                   and r["level"] == "cat" and (r["theme_a_id"], r["theme_b_id"]) == ("a", "b"))
        self.assertEqual((row["titles_total"], row["theme_a_titles"], row["theme_b_titles"],
                          row["shared_titles"], row["shared_per_1000_titles"], row["relative_frequency"]),
                         (4, 2, 2, 1, 250, 1))
        cat_cases = [r for r in self.result["theme_link_cases"] if r["dataset_id"] == "a" and r["level"] == "cat"]
        block_cases = [r for r in self.result["theme_link_cases"] if r["dataset_id"] == "a" and r["level"] == "block"]
        self.assertEqual(len(cat_cases), 3)
        self.assertEqual(len(block_cases), 1)
        self.assertEqual({r["excel_row"] for r in cat_cases}, {6})

    def test_zero_and_undefined_remain_distinct(self):
        # Present a and c in b, but b and c never co-occur: lift really is zero.
        zero = next(r for r in self.result["theme_links"] if r["dataset_id"] == "b"
                    and r["level"] == "cat" and (r["theme_a_id"], r["theme_b_id"]) == ("b", "c"))
        self.assertEqual(zero["relative_frequency"], 0)
        absent = [r for r in self.result["theme_links"] if r["level"] == "cat"
                  and "absent" in (r["theme_a_id"], r["theme_b_id"])]
        self.assertTrue(all(r["relative_frequency"] is None for r in absent))
        empty = [r for r in self.result["document_types"] if r["dataset_id"] == "empty"]
        self.assertTrue(all(r["titles_with_type"] == 0 and r["percent_of_titles"] is None for r in empty))
        self.assertEqual(len(self.result["theme_links"]), 3 * (6 + 3))

    def test_description_counts_and_evidence_reconcile(self):
        for key, fields, value in (("document_types", ("document_type_id",), "titles_with_type"),
                                   ("theme_links", ("level", "theme_a_id", "theme_b_id"), "shared_titles")):
            for row in self.result[key]:
                matching = [r for r in self.result[key + "_by_description"]
                            if all(r[f] == row[f] for f in ("dataset_id", *fields))]
                self.assertEqual(sum(r[value] for r in matching), row[value])
                self.assertEqual(sum(r["titles_total"] for r in matching), row["titles_total"])
        evidence = Counter((r["dataset_id"], r["level"], r["theme_a_id"], r["theme_b_id"])
                           for r in self.result["theme_link_cases"])
        for row in self.result["theme_links"]:
            self.assertEqual(row["shared_titles"], evidence[(row["dataset_id"], row["level"], row["theme_a_id"], row["theme_b_id"])])

    def test_repeated_numbers_and_sheet_names_have_distinct_record_ids(self):
        self.assertEqual(len({r.uid for r in self.records}), len(self.records))

    def test_network_uses_shared_taxonomy_and_escapes_embedded_titles(self):
        self.records[0].title = '</script><img src=x onerror="alert(1)">'
        result = build_comparative_profiles(self.datasets, self.descriptions, self.records, self.categories)
        network = result["theme_network"]
        self.assertEqual(len(network["meta"]["datasets"]), 3)
        self.assertEqual(len(network["meta"]["scopes"]), 7)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "theme_links.html"
            write_theme_network(path, network)
            page = path.read_text()
            self.assertNotIn(self.records[0].title, page)
            payload = page.split('id="tl-data">')[1].split('</script>')[0]
            self.assertEqual(json.loads(payload), network)

    def test_new_tables_export_with_headers_even_without_evidence(self):
        analysis = build_analysis(BASE, self.datasets, self.descriptions, [], self.categories,
                                  {"similarity": {"enabled": False}})
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_csv_reports(root, analysis, [], [])
            for key, columns in REPORT_COLUMNS.items():
                with (root / f"{key}.csv").open(encoding="utf-8-sig", newline="") as stream:
                    reader = csv.DictReader(stream, delimiter=";")
                    self.assertEqual(reader.fieldnames, [key for key, _ in columns])
            write_workbook(root / "comparison.xlsx", self.datasets, self.descriptions, [], [], self.categories, analysis)
            book = load_workbook(root / "comparison.xlsx", read_only=True)
            self.assertIn("Зв’язки тем", book.sheetnames)
            self.assertIn("Типи документів", book.sheetnames)
            book.close()
            write_html_report(root / "report.html", self.datasets, analysis, [], [])
            self.assertIn('href="figures/theme_links.html"', (root / "report.html").read_text())


if __name__ == "__main__":
    unittest.main()
