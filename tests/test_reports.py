from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from src import __version__
from src.models import Dataset
from src.reports import write_manifest


class ReportTests(unittest.TestCase):
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
