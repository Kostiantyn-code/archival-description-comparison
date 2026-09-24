from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from src import __version__
from src.models import Category, Dataset, Description
from src.reports import (
    _description_chronology_panel, _description_colors, _fond_summary,
    _grouped_category_bars, _stacked_description_bars, write_manifest,
)


class ReportTests(unittest.TestCase):
    def test_compared_fonds_are_adjacent_and_inventories_stack(self):
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        left = Dataset("left", "Фонд 230", "ДАМО-ЦДІАК", "", Path("left.xlsx"))
        right = Dataset("right", "Фонд 533", "ЦДІАК", "", Path("right.xlsx"))
        items = [
            Description("left", "left.xlsx", "Опис 2", "ДАМО", "230", "2", "uk", 5, ""),
            Description("left", "left.xlsx", "Опис 3", "ДАМО", "230", "3", "uk", 5, ""),
            Description("left", "left.xlsx", "Опис 1", "ЦДІАК", "356", "1", "uk", 5, ""),
            Description("right", "right.xlsx", "Опис 5", "ЦДІАК", "533", "5", "uk", 5, ""),
        ]
        category = Category("culture", "Культура", "culture", 1, [], [], [], [], [], [], [], [])
        counts = {("left", "Опис 2", "culture"): 6,
                  ("left", "Опис 3", "culture"): 3,
                  ("left", "Опис 1", "culture"): 1,
                  ("right", "Опис 5", "culture"): 4}
        colors = _description_colors(items)
        self.assertNotEqual(colors[("left", "Опис 2")], colors[("left", "Опис 3")])
        self.assertNotEqual(colors[("left", "Опис 1")], colors[("left", "Опис 2")])
        self.assertIn("Фонд 230: Описи 2, 3", _fond_summary(items[:3]))
        self.assertIn("Фонд 356: Опис 1", _fond_summary(items[:3]))

        fig, ax = plt.subplots()
        _grouped_category_bars(ax, [left, right], items, [category], counts, colors,
                               {"left": 20, "right": 8})
        segments = [container.patches[0] for container in ax.containers]
        self.assertEqual([segment.get_y() for segment in segments[:3]],
                         [segments[0].get_y()] * 3)
        self.assertEqual(segments[1].get_x(), 30)
        self.assertEqual(segments[2].get_x(), 45)
        self.assertEqual(segments[3].get_x(), 0)
        self.assertNotEqual(segments[0].get_y(), segments[3].get_y())
        self.assertEqual(sum(segment.get_width() for segment in segments[:3]), 50)
        self.assertEqual(segments[3].get_width(), 50)  # 4/8 equals 10/20.
        plt.close(fig)

        fig, ax = plt.subplots()
        yearly = {(item.dataset_id, item.sheet_name, year): count
                  for item, count in zip(items, (6, 3, 1, 4)) for year in (1850, 1851)}
        _description_chronology_panel(ax, left, items, [1850, 1851], yearly, colors)
        self.assertEqual(len(ax.lines), 3)
        self.assertEqual([line.get_ydata()[0] for line in ax.lines], [6, 3, 1])
        plt.close(fig)

    def test_description_segments_share_one_bar_in_both_orientations(self):
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        first = Description("fund", "fund.xlsx", "Опис 1", "ДАМО", "230", "1", "uk", 5, "ДАМО, оп. 1")
        second = Description("fund", "fund.xlsx", "Опис 2", "ДАМО", "230", "2", "uk", 5, "ДАМО, оп. 2")
        descriptions = [first, second]
        colors = {(item.dataset_id, item.sheet_name): "#366A9A" for item in descriptions}
        values = {(first.dataset_id, first.sheet_name): [6, 3],
                  (second.dataset_id, second.sheet_name): [4, 7]}
        for horizontal in (True, False):
            with self.subTest(horizontal=horizontal):
                fig, ax = plt.subplots()
                _stacked_description_bars(ax, ["A", "B"], descriptions, values, colors,
                                          horizontal=horizontal)
                first_bar = ax.containers[0].patches
                second_bar = ax.containers[1].patches
                for i in range(2):
                    start = first_bar[i].get_x() if horizontal else first_bar[i].get_y()
                    extent = first_bar[i].get_width() if horizontal else first_bar[i].get_height()
                    second_start = second_bar[i].get_x() if horizontal else second_bar[i].get_y()
                    second_extent = second_bar[i].get_width() if horizontal else second_bar[i].get_height()
                    self.assertEqual(start, 0)
                    self.assertEqual(second_start, extent)
                    self.assertEqual(extent + second_extent, 10)
                plt.close(fig)

    def test_manifest_accepts_inputs_outside_project_directory(self):
        with tempfile.TemporaryDirectory() as project_tmp, tempfile.TemporaryDirectory() as input_tmp:
            project_dir = Path(project_tmp)
            input_path = Path(input_tmp) / "external.xlsx"
            input_path.write_bytes(b"test workbook placeholder")
            output_path = project_dir / "run" / "run_manifest.json"
            output_path.parent.mkdir()

            dataset = Dataset(
                id="external",
                label="External dataset",
                short_label="External",
                reference="Archive, fonds 1",
                path=input_path,
            )
            write_manifest(output_path, project_dir, [dataset], "test", [])

            payload = json.loads(output_path.read_text(encoding="utf-8"))
            self.assertEqual(payload["script_version"], __version__)
            self.assertEqual(list(payload["sha256"]), ["external.xlsx"])
            self.assertEqual(len(payload["sha256"]["external.xlsx"]), 64)


if __name__ == "__main__":
    unittest.main()
