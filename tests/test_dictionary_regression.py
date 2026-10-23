"""Frozen pre-cleanup oracle from main 901f451; this is not an accuracy benchmark."""

from __future__ import annotations

from dataclasses import asdict
import hashlib
import json
from pathlib import Path
import tempfile
import unittest

import yaml

from src.classification import classify_title, load_language_dictionaries
from src.models import CATEGORY_RULE_FIELDS

ROOT = Path(__file__).resolve().parents[1]


def fingerprint(value) -> str:
    return hashlib.sha256(
        json.dumps(value, sort_keys=True, ensure_ascii=False).encode()
    ).hexdigest()


class DictionaryRegressionTests(unittest.TestCase):
    def test_rules_and_3900_classification_decisions_match_pre_cleanup_oracle(self):
        expected_rules = {
            "uk": "59bf6c2c7f77af23e8cdf642e76fee45edf55a469f59d77f4d8ccfb50b6ce9ca",
            "ru": "94800adf82edbda1159b620e62f4e73e26bd53337bda731956614292fc0064db",
        }
        decisions = []
        for language in ("uk", "ru"):
            categories, ambiguities, blocks, _ = load_language_dictionaries(
                ROOT, yaml, language
            )
            rules = {
                "categories": [
                    {k: v for k, v in asdict(c).items() if k != "label"}
                    for c in categories
                ],
                "ambiguities": ambiguities,
            }
            self.assertEqual(fingerprint(rules), expected_rules[language])
            titles = set()
            for category in categories:
                for field in CATEGORY_RULE_FIELDS:
                    for item in getattr(category, field):
                        if isinstance(item, str):
                            titles.add(item)
                        else:
                            titles.update(
                                " ".join(item.get("all", []) + ([x] if x else []))
                                for x in item.get("any", [""])
                            )
            for ambiguity in ambiguities:
                titles.add(ambiguity["term"])
                for terms in ambiguity.get("conditional", {}).values():
                    titles.update(ambiguity["term"] + " " + x for x in terms)
                for key, terms in ambiguity.items():
                    if key.startswith("exclude_for_"):
                        titles.update(terms)
            prefix = "Листування про " if language == "uk" else "Переписка о "
            for title in sorted(titles):
                for text in (title, prefix + title):
                    result = classify_title(
                        text, categories, ambiguities, list(blocks), language
                    )
                    decisions.append(
                        {
                            "language": language,
                            "title": text,
                            "categories": sorted(result[0]),
                            "macroblocks": result[2],
                            "scores": result[3],
                            "context": sorted(result[5]),
                            "review": result[8],
                        }
                    )
        self.assertEqual(len(decisions), 3900)
        self.assertEqual(
            fingerprint(decisions),
            "fee115f03aa6dd0636cf7b6d497b536a8e89a5174da6f59563542e6605f62965",
        )

    def test_unknown_language_and_duplicate_category_ids_fail_explicitly(self):
        with self.assertRaisesRegex(ValueError, "Unsupported analysis language"):
            load_language_dictionaries(ROOT, yaml, "hr")
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "config").mkdir()
            (root / "dictionaries").mkdir()
            (root / "config/categories.yaml").write_bytes(
                (ROOT / "config/categories.yaml").read_bytes()
            )
            pack = yaml.safe_load(
                (ROOT / "dictionaries/uk.yaml").read_text(encoding="utf-8")
            )
            pack["categories"].append(pack["categories"][0])
            (root / "dictionaries/uk.yaml").write_text(
                yaml.safe_dump(pack, allow_unicode=True), encoding="utf-8"
            )
            with self.assertRaisesRegex(ValueError, "category IDs"):
                load_language_dictionaries(root, yaml, "uk")
