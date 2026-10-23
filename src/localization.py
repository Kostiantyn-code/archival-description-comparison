"""English/Ukrainian presentation without translating archival source text."""

from __future__ import annotations

import html
import json
import re
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

CATALOG = Path(__file__).resolve().parent.parent / "config" / "localization.json"


@lru_cache(maxsize=1)
def catalog() -> dict:
    return json.loads(CATALOG.read_text(encoding="utf-8"))


@lru_cache(maxsize=1)
def ukrainian_captions() -> dict[str, str]:
    return catalog()["uk_ui"]


@dataclass(frozen=True)
class Visualization:
    language: str = "en"
    bilingual_labels: bool = False
    label_overrides: dict = field(default_factory=dict)

    @classmethod
    def from_config(cls, config: dict) -> "Visualization":
        options = config.get("visualization", {})
        if not isinstance(options, dict):
            raise ValueError("visualization: expected a mapping")
        language = options.get("language", "en")
        bilingual = options.get("bilingual_labels", False)
        overrides = options.get("label_overrides", {})
        if language not in ("uk", "en"):
            raise ValueError("visualization.language: expected uk or en")
        if not isinstance(bilingual, bool):
            raise ValueError("visualization.bilingual_labels: expected true or false")
        if not isinstance(overrides, dict) or any(
            not isinstance(group, str)
            or not isinstance(values, dict)
            or any(
                not isinstance(key, str) or not isinstance(value, str)
                for key, value in values.items()
            )
            for group, values in overrides.items()
        ):
            raise ValueError(
                "visualization.label_overrides: expected groups of string labels"
            )
        return cls(language, bilingual, overrides)

    @property
    def bilingual(self) -> bool:
        return self.language == "en" and self.bilingual_labels

    def text(self, original: str) -> str:
        """Translate a trusted English caption, never arbitrary source values."""
        return (
            ukrainian_captions().get(original, original)
            if self.language == "uk"
            else original
        )

    def label(self, group: str, key: str, original: str) -> str:
        ukrainian = catalog()["uk_labels"].get(group, {}).get(key, original)
        if self.language == "uk":
            return ukrainian
        translated = self.label_overrides.get(group, {}).get(
            key, catalog().get(group, {}).get(key, original)
        )
        if self.bilingual and translated != ukrainian:
            return f"{translated} ({ukrainian})"
        return translated

    def dataset(self, key: str, original: str, *, short: bool = False) -> str:
        group = "dataset_short" if short else "datasets"
        if short and key not in self.label_overrides.get(group, {}):
            group = "datasets"
        return self.label(group, key, original)

    def inventory_name(self, original: str) -> str:
        if self.language == "en" and (
            match := re.fullmatch(r"Опис\s+(\d+)", original, re.I)
        ):
            translated = f"Inventory {match[1]}"
            return f"{translated} ({original})" if self.bilingual else translated
        return original

    def reference(self, archive: str, fond: str, inventory: str) -> str:
        archive = self.label("archives", archive, archive)
        if self.language == "uk":
            return (
                archive
                + (f", ф. {fond}" if fond else "")
                + (f": опис {inventory}" if inventory else "")
            )
        return (
            archive
            + (f", fond {fond}" if fond else "")
            + (f": inventory {inventory}" if inventory else "")
        )

    def inventory_reference(
        self, key: str, archive: str, fond: str, inventory: str, original: str
    ) -> str:
        if key in self.label_overrides.get("descriptions", {}):
            return self.label("descriptions", key, original)
        return self.reference(archive, fond, inventory) or self.inventory_name(original)

    def rows(self, rows: list[dict]) -> list[dict]:
        """Localize only known generated fields; preserve source text verbatim."""
        displayed = []
        for source in rows:
            row = dict(source)
            for field, id_field, group in (
                ("category", "category_id", "categories"),
                ("document_type", "document_type_id", "document_types"),
                (
                    "theme_a",
                    "theme_a_id",
                    "categories" if row.get("level") == "cat" else "macroblocks",
                ),
                (
                    "theme_b",
                    "theme_b_id",
                    "categories" if row.get("level") == "cat" else "macroblocks",
                ),
            ):
                if field in row and id_field in row:
                    row[field] = self.label(group, row[id_field], row[field])
            for field in ("type", "classification_status"):
                if field in row:
                    row[field] = self.text(row[field])
            if "dataset" in row and "dataset_id" in row:
                row["dataset"] = self.dataset(row["dataset_id"], row["dataset"])
            displayed.append(row)
        return displayed

    def template(self, source: str) -> str:
        """Render explicit UI markers before embedding any archival data.

        Markers keep JS identifiers and CSS outside the translation path.
        Captions are escaped for their HTML or JavaScript context.
        """

        def javascript(value: str) -> str:
            return (
                json.dumps(value, ensure_ascii=False)[1:-1]
                .replace("'", r"\u0027")
                .replace("`", r"\u0060")
                .replace("$", r"\u0024")
                .replace("<", r"\u003c")
                .replace("\u2028", r"\u2028")
                .replace("\u2029", r"\u2029")
            )

        parts = re.split(r"(<script\b[^>]*>.*?</script>)", source, flags=re.S | re.I)
        for index, part in enumerate(parts):
            escape = javascript if part.lower().startswith("<script") else html.escape
            parts[index] = re.sub(
                r"{{ui:([^{}]+)}}", lambda match: escape(self.text(match[1])), part
            )
        page = "".join(parts)
        return page.replace('lang="en"', f'lang="{self.language}"').replace(
            "'en-GB'", "'uk-UA'" if self.language == "uk" else "'en-GB'"
        )

    def metadata(self) -> dict:
        return {
            "language": self.language,
            "bilingual_labels": self.bilingual_labels,
            "label_overrides": self.label_overrides,
        }
