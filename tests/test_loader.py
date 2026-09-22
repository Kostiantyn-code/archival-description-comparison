from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from src.analysis import overview_rows
from src.loader import load_all


class LoaderTests(unittest.TestCase):
    def _book(self, path: Path, archive: str, fond: str, inventory: str, shifted: bool = False) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = f"Опис {inventory}"
        sheet.append(["Архів", archive])
        sheet.append(["Фонд", fond])
        sheet.append(["Опис", inventory])
        sheet.append(["№ справи", "Заголовок справи", "Крайні дати", "Кількість аркушів", "Примітки"])
        if shifted:
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


if __name__ == "__main__":
    unittest.main()
