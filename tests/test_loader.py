from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from src.analysis import overview_rows
from src.loader import load_all


class LoaderTests(unittest.TestCase):
    def _book(self, path: Path, archive: str, fond: str, inventory: str,
              shifted: bool = False, language: str | None = None,
              titles: list[str] | None = None) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = f"Опис {inventory}"
        sheet.append(["Архів", archive])
        sheet.append(["Фонд", fond])
        sheet.append(["Опис", inventory])
        if language is not None:
            sheet.append(["Мова", language])
        sheet.append(["№ справи", "Заголовок справи", "Крайні дати", "Кількість аркушів", "Примітки"])
        if titles is not None:
            for number, title in enumerate(titles, 1):
                sheet.append([number, title, "1850", 10, None])
        elif shifted:
            sheet.append([1, "1918 рік", "1 січня 1918", 10, None])
            sheet.append([None, "Листування про відкриття школи", None, None, None])
        else:
            sheet.append([1, "Справа про міську лікарню", "1918", 12, None])
        workbook.save(path)

    def test_files_are_datasets_and_sheets_are_descriptions(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._book(root / "one.xlsx", "Архів А", "1", "1")
            self._book(root / "two.xlsx", "Архів Б", "2", "1")
            datasets, descriptions, records, _ = load_all(
                root, {"datasets": []}, {"minimum_year": 1700, "maximum_year": 2026}
            )
            self.assertEqual(len(datasets), 2)
            self.assertEqual(len(descriptions), 2)
            self.assertEqual(len(records), 2)
            rows = overview_rows(datasets, descriptions, records)
            self.assertEqual([row["cases"] for row in rows], [1, 1])

    def test_shifted_year_row_is_joined_with_following_title(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._book(root / "one.xlsx", "Архів А", "1", "1", shifted=True)
            self._book(root / "two.xlsx", "Архів Б", "2", "1")
            _, _, records, issues = load_all(
                root, {"datasets": []}, {"minimum_year": 1700, "maximum_year": 2026}
            )
            record = next(item for item in records if item.file_name == "one.xlsx")
            self.assertEqual(record.case_id, "1")
            self.assertEqual(record.title, "Листування про відкриття школи")
            self.assertEqual(record.start_year, 1918)
            self.assertEqual(record.pages, 10)
            self.assertTrue(any(issue.field == "Зміщений рядок" for issue in issues))

    def test_one_dataset_keeps_distinct_archives_and_fonds(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._book(root / "combined.xlsx", "ДАМО", "230", "1")
            self._book(root / "other.xlsx", "ДАМО", "229", "1")
            book = load_workbook(root / "combined.xlsx")
            sheet = book.copy_worksheet(book.active)
            sheet.title = "ЦДІАК"
            sheet.cell(1, 2, "ЦДІАК")
            sheet.cell(2, 2, "356")
            book.save(root / "combined.xlsx")
            book.close()
            datasets, descriptions, records, _ = load_all(root, {"datasets": []}, {})
            combined = next(d for d in datasets if d.path.name == "combined.xlsx")
            self.assertEqual(len(datasets), 2)
            parts = [d for d in descriptions if d.dataset_id == combined.id]
            self.assertEqual([(d.archive, d.fond, d.inventory) for d in parts],
                             [("ДАМО", "230", "1"), ("ЦДІАК", "356", "1")])
            self.assertIn("356", combined.reference)
            self.assertIn("230", combined.reference)
            active = [r for r in records if r.dataset_id == combined.id and r.analyzable]
            self.assertEqual(len(active), 2)
            self.assertEqual(len({r.uid for r in records}), 3)

    def test_all_titles_use_inferred_sheet_language(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._book(root / "russian.xlsx", "ЦДІАК України", "356", "1", titles=[
                "Дело о нарушении карантинного режима",
                "Переписка о строительстве больницы",
                "Журнал исходящих секретных документов",
            ])
            self._book(root / "ukrainian.xlsx", "ДАМО", "230", "1", titles=[
                "Справа про відкриття школи та навчання дітей",
                "Журнал вхідних документів",
            ])
            _, descriptions, records, _ = load_all(root, {"datasets": []}, {})
            self.assertEqual([(d.sheet_name, d.language) for d in descriptions],
                             [("Опис 1", "ru"), ("Опис 1", "uk")])
            self.assertEqual([r.language for r in records if r.file_name == "russian.xlsx"],
                             ["ru", "ru", "ru"])
            self.assertEqual([r.language for r in records if r.file_name == "ukrainian.xlsx"],
                             ["uk", "uk"])

    def test_explicit_language_applies_to_entire_sheet(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._book(root / "one.xlsx", "ЦДІАК України", "356", "1", language="ru",
                       titles=["Справа про відкриття школи та навчання дітей"])
            self._book(root / "two.xlsx", "ДАМО", "230", "1")
            _, descriptions, records, _ = load_all(root, {"datasets": []}, {})
            self.assertEqual(next(d.language for d in descriptions if d.file_name == "one.xlsx"), "ru")
            self.assertEqual(next(r.language for r in records if r.file_name == "one.xlsx"), "ru")


if __name__ == "__main__":
    unittest.main()
