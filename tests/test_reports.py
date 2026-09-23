from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from src import __version__
from src.models import Dataset, Description
from src.reports import _stacked_description_bars, write_manifest


class ReportTests(unittest.TestCase):
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
