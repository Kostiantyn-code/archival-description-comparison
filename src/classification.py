"""Тематична класифікація заголовків справ на основі словників."""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from .models import CATEGORY_RULE_FIELDS, Category, Record
from .text_matching import (
    MATCH_LANGUAGE,
    item_matcher,
    mask_item,
    stem_tokens,
)


def _load_yaml(path: Path, yaml_module) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as stream:
        data = yaml_module.safe_load(stream)
    if not isinstance(data, dict):
        raise ValueError(f"Корінь YAML має бути словником: {path}")
    return data


def load_language_dictionaries(
    base_dir: Path,
    yaml_module,
    language: str,
) -> tuple[list[Category], list[dict[str, Any]], dict[str, str], str]:
    index_path = base_dir / "config" / "categories.yaml"
    index = _load_yaml(index_path, yaml_module)
    categories: list[Category] = []
    for item in index.get("categories", []):
        if not item.get("enabled", True):
            continue
        dictionary_path = base_dir / item["dictionaries"][language]
        data = _load_yaml(dictionary_path, yaml_module)
        categories.append(Category(
            id=str(data["id"]),
            label=str(item["label"]),
            macroblock=str(data["macroblock"]),
            minimum_score=int(data.get("minimum_score", index.get("classification", {}).get("minimum_score", 3))),
            **{field: list(data.get(field, [])) for field in CATEGORY_RULE_FIELDS},
        ))
    ambiguity_path = base_dir / "dictionaries" / ("ru" if language == "ru" else "") / "auxiliary" / "ambiguities.yaml"
    ambiguities = _load_yaml(ambiguity_path, yaml_module).get("rules", []) if ambiguity_path.exists() else []
    macroblocks = {
        key: str(value.get("label", key))
        for key, value in index.get("macroblocks", {}).items()
    }
    return categories, ambiguities, macroblocks, str(index.get("dictionary_version", "невідома"))


def _global_masks(tokens: list[str], category_id: str, ambiguities: list[dict[str, Any]]) -> list[str]:
    result = list(tokens)
    for rule in ambiguities:
        term = str(rule.get("term", ""))
        if category_id in rule.get("do_not_assign", []):
            result = mask_item(result, term)
        for phrase in rule.get(f"exclude_for_{category_id}", []):
            result = mask_item(result, str(phrase))
    return result


def _context_rule(matches, rule: dict[str, Any]) -> bool:
    all_items = [str(value) for value in rule.get("all", [])]
    any_items = [str(value) for value in rule.get("any", [])]
    if all_items and not all(matches(item) for item in all_items):
        return False
    if any_items and not any(matches(item) for item in any_items):
        return False
    return bool(all_items or any_items)


def classify_title(
    title: str,
    categories: list[Category],
    ambiguities: list[dict[str, Any]],
    macroblock_order: list[str],
    language: str,
):
    language_token = MATCH_LANGUAGE.set(language)
    try:
        original_tokens = stem_tokens(title)
        matched_ids: list[str] = []
        labels: list[str] = []
        scores: dict[str, int] = {}
        evidence: dict[str, list[str]] = {}
        context_ids: list[str] = []
        context_labels: list[str] = []
        context_evidence: dict[str, list[str]] = {}
        review_flags: set[str] = set()

        for category in categories:
            tokens = _global_masks(original_tokens, category.id, ambiguities)
            for phrase in category.exclude_phrases:
                tokens = mask_item(tokens, phrase)
            matches = item_matcher(tokens)

            weak_evidence: list[str] = []
            for items, label in (
                (category.context_only_phrases, "фраза"),
                (category.context_only_terms, "термін"),
                (category.contextual_terms, "контекст"),
            ):
                weak_evidence.extend(f"{label}: {item}" for item in items if matches(item))
            if weak_evidence:
                context_ids.append(category.id)
                context_labels.append(category.label)
                context_evidence[category.id] = weak_evidence

            score = 0
            category_evidence: list[str] = []
            for items, weight, label in (
                (category.strong_phrases, 4, "фраза"),
                (category.strong_terms, 3, "термін"),
                (category.contextual_terms, 1, "контекст"),
            ):
                for item in items:
                    if matches(item):
                        score += weight
                        category_evidence.append(f"{label}: {item}")
            for rule in category.context_rules:
                if _context_rule(matches, rule):
                    score += int(rule.get("score", 3))
                    readable = " + ".join(
                        str(value)
                        for key in ("all", "any")
                        for value in rule.get(key, [])
                        if matches(str(value))
                    )
                    category_evidence.append(f"правило: {readable}")

            for ambiguity in ambiguities:
                term = str(ambiguity.get("term", ""))
                contexts = ambiguity.get("conditional", {}).get(category.id, [])
                if term and contexts and matches(term) and any(matches(str(item)) for item in contexts):
                    score += 3
                    category_evidence.append(f"контекст неоднозначного терміна: {term}")

            has_subject = any(not item.startswith("контекст:") for item in category_evidence)
            if score >= category.minimum_score and (language != "ru" or has_subject):
                matched_ids.append(category.id)
                labels.append(category.label)
                scores[category.id] = score
                evidence[category.id] = category_evidence
            elif score >= max(1, category.minimum_score - 1):
                review_flags.update(
                    f"{category.id}: {term}"
                    for term in category.review_terms
                    if matches(term)
                )

        ordered = sorted(zip(matched_ids, labels), key=lambda pair: (-scores[pair[0]], pair[1]))
        matched_ids = [item[0] for item in ordered]
        labels = [item[1] for item in ordered]
        category_by_id = {category.id: category for category in categories}
        present_macroblocks = {category_by_id[item].macroblock for item in matched_ids}
        macroblocks = [item for item in macroblock_order if item in present_macroblocks]
        return (
            matched_ids,
            labels,
            macroblocks,
            scores,
            evidence,
            context_ids,
            context_labels,
            context_evidence,
            sorted(review_flags),
        )
    finally:
        MATCH_LANGUAGE.reset(language_token)


def classify_records(
    records: list[Record],
    language_categories: dict[str, list[Category]],
    language_ambiguities: dict[str, list[dict[str, Any]]],
    macroblock_order: list[str],
    progress=None,
) -> None:
    active = [record for record in records if record.analyzable]
    total = len(active)
    for index, record in enumerate(active, start=1):
        categories = language_categories.get(record.language, language_categories["uk"])
        ambiguities = language_ambiguities.get(record.language, language_ambiguities["uk"])
        (
            record.categories,
            record.category_labels,
            record.macroblocks,
            record.scores,
            record.evidence,
            record.context_categories,
            record.context_category_labels,
            record.context_evidence,
            record.review_flags,
        ) = classify_title(
            record.title,
            categories,
            ambiguities,
            macroblock_order,
            record.language,
        )
        if progress and (index == total or index % max(1, total // 100) == 0):
            progress(index, total)
