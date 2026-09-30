"""Shared YAML loading keeps the existing entry points and error policy."""
import tempfile
import unittest
from pathlib import Path

import yaml

from src.configuration import safe_load_yaml
from src.loader import safe_load_yaml as loader_yaml
from src.classification import _load_yaml as classification_yaml


class ConfigurationTests(unittest.TestCase):
    def test_config_and_dictionary_readers_accept_the_same_mapping(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "config.yaml"
            path.write_text('name: "Миколаїв"\nrules: [школа, церква]\n', encoding="utf-8")
            for reader in (safe_load_yaml, loader_yaml, classification_yaml):
                self.assertEqual(reader(path, yaml), {"name": "Миколаїв", "rules": ["школа", "церква"]})

    def test_non_mapping_and_malformed_yaml_keep_their_errors(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "config.yaml"
            for reader in (safe_load_yaml, loader_yaml, classification_yaml):
                for content in ("", "- rule", "scalar", "null"):
                    path.write_text(content, encoding="utf-8")
                    with self.assertRaises(ValueError) as error:
                        reader(path, yaml)
                    self.assertEqual(str(error.exception), f"Корінь YAML має бути словником: {path}")
                path.write_text("rules: [", encoding="utf-8")
                with self.assertRaises(yaml.YAMLError):
                    reader(path, yaml)


if __name__ == "__main__":
    unittest.main()
