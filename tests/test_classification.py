"""Subject rules and the scope of explicitly named personnel sections."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import yaml
from openpyxl import Workbook

from src.classification import classify_records, classify_title, load_language_dictionaries
from src.loader import load_all


ROOT = Path(__file__).resolve().parents[1]


class ClassificationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.categories = {}
        cls.ambiguities = {}
        for language in ("uk", "ru"):
            cls.categories[language], cls.ambiguities[language], blocks, _ = (
                load_language_dictionaries(ROOT, yaml, language)
            )
        cls.blocks = list(blocks)

    def categories_for(self, title: str, language: str = "uk") -> list[str]:
        return classify_title(
            title, self.categories[language], self.ambiguities[language],
            self.blocks, language,
        )[0]

    def test_residence_and_port_charges_require_subject_phrase(self):
        self.assertIn("population_and_society", self.categories_for("Право на проживання євреїв"))
        self.assertIn("population_and_society", self.categories_for("Право проживання"))
        self.assertIn("economy", self.categories_for("Перегляд портових зборів"))
        self.assertIn("economy", self.categories_for("Портовий збір"))
        self.assertNotIn("economy", self.categories_for("Ремонт портових споруд"))

    def test_personnel_terms_and_russian_port_charges(self):
        self.assertIn("personal_files", self.categories_for("Особові справи лікарів"))
        self.assertIn("personal_files", self.categories_for("Списки учнів гімназії"))
        self.assertIn("personal_files", self.categories_for("Личные дела заключенных", "ru"))
        self.assertIn("education", self.categories_for("Списки учеников гимназии", "ru"))
        self.assertIn("personal_files", self.categories_for("Список учеников гимназии", "ru"))
        self.assertIn("economy", self.categories_for("Пересмотр портовых сборов", "ru"))
        self.assertNotIn("personal_files", self.categories_for("Прізвище Іваненка"))

    def test_personnel_headings_survive_letters_and_years_but_stop_at_new_section(self):
        with tempfile.TemporaryDirectory() as tmp:
            input_dir = Path(tmp)
            book = Workbook()
            sheet = book.active
            sheet.append(["Архів", "ДАМО"])
            sheet.append(["Фонд", "229"])
            sheet.append(["Опис", "1"])
            sheet.append(["№ справи", "Заголовок справи", "Крайні дати"])
            sheet.append([None, "Особові справи ув’язнених"])
            sheet.append([None, "«А»"])
            sheet.append([1, "Андрієв Іван (Андреев Иван)", "1902"])
            sheet.append([None, "1903 рік"])
            sheet.append([2, "Бойко Петро (Бойко Петр)", None])
            sheet.append([None, "Поточна кореспонденція"])
            sheet.append([3, "Василенко Олег (Василенко Олег)", "1904"])
            sheet.append([None, "Списки учнів гімназії"])
            sheet.append([4, "Гнатюк Марія (Гнатюк Мария)", "1905"])
            sheet.append([5, "Опис № 1 фонду № 229 за 1900–1905 роки"])
            book.save(input_dir / "one.xlsx")
            book.save(input_dir / "two.xlsx")

            _, _, records, _ = load_all(
                input_dir, {"datasets": []},
                {"minimum_year": 1700, "maximum_year": 2026},
            )
            classify_records(records, self.categories, self.ambiguities, self.blocks)
            first = [record for record in records if record.file_name == "one.xlsx"]
            self.assertEqual([record.thematic_section for record in first[:4]], [
                "Особові справи ув’язнених", "Особові справи ув’язнених", "",
                "Списки учнів гімназії",
            ])
            self.assertFalse(first[4].analyzable)
            self.assertEqual(first[1].start_year, 1903)
            self.assertEqual(set(first[0].categories), {"personal_files", "law_and_police"})
            self.assertEqual(set(first[1].categories), {"personal_files", "law_and_police"})
            self.assertFalse(first[2].categories)
            self.assertEqual(set(first[3].categories), {"personal_files", "education"})
            self.assertEqual(first[0].evidence["law_and_police"], [
                "розділ опису: Особові справи ув’язнених"
            ])


if __name__ == "__main__":
    unittest.main()
