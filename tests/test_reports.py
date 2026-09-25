from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from src import __version__
from src.models import Category, Dataset, Description
from src.reports import (
    _category_figure, _description_chronology_panel, _description_colors,
    _fond_summary, _stacked_description_bars, write_manifest,
)


class ReportTests(unittest.TestCase):
    def test_compared_fonds_are_adjacent_and_inventories_stack(self):
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        left = Dataset("left", "Фонд 230", "ДАМО-ЦДІАК", "", Path("left.xlsx"))
        right = Dataset("right", "Фонд 229", "ДАМО", "", Path("right.xlsx"))
        items = [
            Description("left", "left.xlsx", "Опис 2", "ДАМО", "230", "2", "uk", 5, ""),
            Description("left", "left.xlsx", "Опис 3", "ДАМО", "230", "3", "uk", 5, ""),
            Description("left", "left.xlsx", "Опис 1", "ЦДІАК", "356", "1", "uk", 5, ""),
            Description("right", "right.xlsx", "Опис 1", "ДАМО", "229", "1", "uk", 5, ""),
            Description("right", "right.xlsx", "Опис 2", "ДАМО", "229", "2", "uk", 5, ""),
        ]
        category = Category("culture", "Культура", "culture", 1, [], [], [], [], [], [], [], [])
        counts = {("left", "Опис 2", "culture"): 6,
                  ("left", "Опис 3", "culture"): 3,
                  ("left", "Опис 1", "culture"): 1,
                  ("right", "Опис 1", "culture"): 2,
                  ("right", "Опис 2", "culture"): 2}
        colors = _description_colors(items)
        self.assertNotEqual(colors[("left", "Опис 2")], colors[("left", "Опис 3")])
        self.assertNotEqual(colors[("left", "Опис 1")], colors[("left", "Опис 2")])
        self.assertIn("Фонд 230: Описи 2, 3", _fond_summary(items[:3]))
        self.assertIn("Фонд 356: Опис 1", _fond_summary(items[:3]))

        fig = _category_figure([left, right], items, [category], counts, colors,
                               {"left": 20, "right": 8})
        self.assertEqual(len(fig.axes), 2)
        left_segments = [container.patches[0] for container in fig.axes[0].containers]
        right_segments = [container.patches[0] for container in fig.axes[1].containers]
        self.assertEqual([segment.get_y() for segment in left_segments],
                         [right_segments[0].get_y()] * 3)
        self.assertEqual([segment.get_x() for segment in left_segments], [0, 30, 45])
        self.assertEqual([segment.get_x() for segment in right_segments], [0, 25])
        self.assertEqual(sum(segment.get_width() for segment in left_segments), 50)
        self.assertEqual(sum(segment.get_width() for segment in right_segments), 50)
        self.assertEqual(fig.axes[0].get_xlim(), fig.axes[1].get_xlim())
        self.assertTrue(fig.axes[0].yaxis_inverted())
        self.assertEqual([text.get_text() for text in fig.legends[0].get_texts()], [
            "ДАМО, ф. 230, оп. 2, 3\nЦДІАК, ф. 356, оп. 1",
            "ДАМО, ф. 229, оп. 1, 2",
        ])
        plt.close(fig)

        fig, ax = plt.subplots()
        yearly = {(item.dataset_id, item.sheet_name, year): count
        for item, count in zip(items, (6, 3, 1, 2, 2)) for year in (1850, 1851)}
        description_totals = {("left", "Опис 2"): 12,
                              ("left", "Опис 3"): 6,
                              ("left", "Опис 1"): 4,
                              ("right", "Опис 1"): 4,
                              ("right", "Опис 2"): 8}
        _description_chronology_panel(
            ax, left, items, [1850, 1851], yearly, colors, description_totals
        )
        self.assertEqual(len(ax.lines), 3)
        self.assertEqual([line.get_ydata()[0] for line in ax.lines], [50, 50, 25])
        self.assertIn("100% для кожного опису окремо", ax.get_title(loc="left"))
        self.assertEqual(ax.get_ylabel(), "Частка справ опису, %")
        self.assertEqual(ax.get_ylim(), (0, 100))
        self.assertEqual(len(ax.child_axes), 0)
        plt.close(fig)

        fig, ax = plt.subplots()
        _description_chronology_panel(
            ax, right, items, [1850, 1851], yearly, colors, description_totals
        )
        self.assertEqual([line.get_ydata()[0] for line in ax.lines], [50, 25])
        plt.close(fig)

        fig, ax = plt.subplots()
        uneven = {(item.dataset_id, item.sheet_name, year): count
                  for item, count in zip(items[:3], (713, 14, 60))
                  for year in (1850, 1851)}
        uneven_totals = {("left", "Опис 2"): 1200,
                         ("left", "Опис 3"): 20,
                         ("left", "Опис 1"): 120}
        _description_chronology_panel(
            ax, left, items, [1850, 1851], uneven, colors, uneven_totals
        )
        self.assertEqual([round(float(line.get_ydata()[0]), 2) for line in ax.lines],
                         [59.42, 70, 50])
        self.assertEqual(len(ax.child_axes), 0)
        self.assertEqual([line.get_color() for line in ax.lines],
                         [colors[(item.dataset_id, item.sheet_name)] for item in items[:3]])
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
