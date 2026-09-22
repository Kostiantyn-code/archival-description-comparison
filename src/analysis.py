"""Порівняльні показники, лексика та пошук подібних справ."""
from __future__ import annotations

import math
import statistics
from collections import Counter, defaultdict
from itertools import combinations
from pathlib import Path
from typing import Any

from .models import Category, Dataset, Description, Record
from .text_matching import MATCH_LANGUAGE, light_stem_word, tokenize


def _safe_percent(part: int, whole: int) -> float:
    return part / whole * 100 if whole else 0.0


def _year_values(record: Record, chronology: dict[str, Any]) -> list[int]:
    if record.start_year is None:
        return []
    if chronology.get("expansion_mode", "span") != "span" or record.end_year is None:
        return [record.start_year]
    maximum_span = int(chronology.get("maximum_span_years", 100))
    if record.end_year < record.start_year or record.end_year - record.start_year > maximum_span:
        return [record.start_year]
    return list(range(record.start_year, record.end_year + 1))


def overview_rows(
    datasets: list[Dataset],
    descriptions: list[Description],
    records: list[Record],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for dataset in datasets:
        subset = [record for record in records if record.dataset_id == dataset.id]
        cases = [record for record in subset if record.status == "case"]
        active = [record for record in subset if record.analyzable]
        dated = [record for record in active if record.start_year is not None]
        page_values = [record.pages for record in active if record.pages is not None]
        subject_classified = sum(bool(record.categories) for record in active)
        context_only = sum(
            not record.categories and bool(record.context_categories)
            for record in active
        )
        unclassified = len(active) - subject_classified - context_only
        rows.append({
            "dataset_id": dataset.id,
            "dataset": dataset.label,
            "short_label": dataset.short_label,
            "reference": dataset.reference,
            "files": dataset.path.name,
            "descriptions": sum(item.dataset_id == dataset.id for item in descriptions),
            "cases": len(cases),
            "analyzable_titles": len(active),
            "withdrawn": sum(record.status == "withdrawn" for record in subset),
            "start_year": min((record.start_year for record in dated), default=None),
            "end_year": max((record.end_year or record.start_year for record in dated), default=None),
            "pages_total": sum(page_values),
            "pages_mean": round(statistics.mean(page_values), 2) if page_values else None,
            "pages_median": round(statistics.median(page_values), 2) if page_values else None,
            "missing_titles": sum(not record.title.strip() for record in cases),
            "missing_dates": sum(not record.dates_raw for record in active),
            "missing_pages": sum(record.pages is None for record in active),
            "classified_titles": subject_classified,
            "classification_coverage_percent": round(
                _safe_percent(subject_classified, len(active)), 2
            ),
            "context_only_titles": context_only,
            "context_only_percent": round(
                _safe_percent(context_only, len(active)), 2
            ),
            "unclassified_titles": unclassified,
            "unclassified_percent": round(
                _safe_percent(unclassified, len(active)), 2
            ),
            "category_assignments": sum(len(record.categories) for record in active),
        })
    return rows


def classification_coverage_rows(
    datasets: list[Dataset],
    records: list[Record],
) -> list[dict[str, Any]]:
    """Summarize mutually exclusive classification states per logical dataset."""
    rows: list[dict[str, Any]] = []
    for dataset in datasets:
        active = [
            record for record in records
            if record.dataset_id == dataset.id and record.analyzable
        ]
        subject_classified = sum(bool(record.categories) for record in active)
        context_only = sum(
            not record.categories and bool(record.context_categories)
            for record in active
        )
        unclassified = len(active) - subject_classified - context_only
        rows.append({
            "dataset_id": dataset.id,
            "dataset": dataset.label,
            "short_label": dataset.short_label,
            "reference": dataset.reference,
            "analyzable_titles": len(active),
            "subject_classified": subject_classified,
            "subject_classified_percent": round(
                _safe_percent(subject_classified, len(active)), 2
            ),
            "context_only": context_only,
            "context_only_percent": round(
                _safe_percent(context_only, len(active)), 2
            ),
            "unclassified": unclassified,
            "unclassified_percent": round(
                _safe_percent(unclassified, len(active)), 2
            ),
        })
    return rows


def topic_unclassified_rows(records: list[Record]) -> list[dict[str, Any]]:
    """Return analyzable cases without a subject category for dictionary review."""
    rows: list[dict[str, Any]] = []
    for record in records:
        if not record.analyzable or record.categories:
            continue
        rows.append({
            "record_uid": record.uid,
            "dataset_id": record.dataset_id,
            "dataset": record.dataset_label,
            "file_name": record.file_name,
            "sheet_name": record.sheet_name,
            "excel_row": record.excel_row,
            "archive": record.archive,
            "fond": record.fond,
            "inventory": record.inventory,
            "case_id": record.case_id,
            "title": record.title,
            "dates_raw": record.dates_raw,
            "classification_status": (
                "Лише контекст" if record.context_categories
                else "Не класифіковано"
            ),
            "context_categories": " | ".join(record.context_category_labels),
            "review_flags": " | ".join(record.review_flags),
        })
    return rows


def description_rows(
    descriptions: list[Description],
    records: list[Record],
    dataset_by_id: dict[str, Dataset],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for description in descriptions:
        subset = [
            record for record in records
            if record.dataset_id == description.dataset_id and record.sheet_name == description.sheet_name
        ]
        cases = [record for record in subset if record.status == "case"]
        active = [record for record in subset if record.analyzable]
        rows.append({
            "dataset_id": description.dataset_id,
            "dataset": dataset_by_id[description.dataset_id].label,
            "file_name": description.file_name,
            "sheet_name": description.sheet_name,
            "archive": description.archive,
            "fond": description.fond,
            "inventory": description.inventory,
            "reference": description.reference,
            "header_row": description.header_row,
            "cases": len(cases),
            "analyzable_titles": len(active),
            "withdrawn": sum(record.status == "withdrawn" for record in subset),
            "classified_titles": sum(bool(record.categories) for record in active),
            "missing_dates": sum(not record.dates_raw for record in active),
            "missing_pages": sum(record.pages is None for record in active),
        })
    return rows


def chronology_rows(
    datasets: list[Dataset],
    records: list[Record],
    chronology: dict[str, Any],
) -> list[dict[str, Any]]:
    counts: dict[str, Counter[int]] = {dataset.id: Counter() for dataset in datasets}
    for record in records:
        if not record.analyzable:
            continue
        counts[record.dataset_id].update(_year_values(record, chronology))
    years = sorted({year for counter in counts.values() for year in counter})
    rows: list[dict[str, Any]] = []
    for year in years:
        row: dict[str, Any] = {"year": year}
        for dataset in datasets:
            row[dataset.id] = counts[dataset.id][year]
        rows.append(row)
    return rows


def category_rows(
    datasets: list[Dataset],
    records: list[Record],
    categories: list[Category],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    category_by_id = {category.id: category for category in categories}
    for dataset in datasets:
        active = [record for record in records if record.dataset_id == dataset.id and record.analyzable]
        counts = Counter(category_id for record in active for category_id in record.categories)
        for category in categories:
            rows.append({
                "dataset_id": dataset.id,
                "dataset": dataset.label,
                "category_id": category.id,
                "category": category.label,
                "macroblock": category_by_id[category.id].macroblock,
                "cases": counts[category.id],
                "dataset_titles_total": len(active),
                "percent_of_titles": round(_safe_percent(counts[category.id], len(active)), 2),
            })
    return rows


def _read_stopwords(config_dir: Path) -> set[str]:
    result: set[str] = set()
    for language in ("uk", "ru"):
        path = config_dir / f"stopwords_{language}.txt"
        words = {
            line.strip().casefold()
            for line in path.read_text(encoding="utf-8").splitlines()
            if line.strip() and not line.lstrip().startswith("#")
        }
        language_token = MATCH_LANGUAGE.set(language)
        try:
            result.update(light_stem_word(word) for word in words)
        finally:
            MATCH_LANGUAGE.reset(language_token)
    return result


def _document_terms(record: Record, stopword_stems: set[str]) -> tuple[set[str], dict[str, Counter[str]]]:
    language_token = MATCH_LANGUAGE.set(record.language)
    stems: set[str] = set()
    surfaces: dict[str, Counter[str]] = defaultdict(Counter)
    try:
        for surface in tokenize(record.title):
            surface = surface.casefold()
            if surface.isdigit() or len(surface) < 3:
                continue
            stem = light_stem_word(surface)
            if not stem or stem in stopword_stems or len(stem) < 3:
                continue
            stems.add(stem)
            surfaces[stem][surface] += 1
    finally:
        MATCH_LANGUAGE.reset(language_token)
    return stems, surfaces


def vocabulary_tables(
    datasets: list[Dataset],
    records: list[Record],
    config_dir: Path,
    vocabulary_config: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    stopwords = _read_stopwords(config_dir)
    document_frequency: dict[str, Counter[str]] = {dataset.id: Counter() for dataset in datasets}
    surfaces: dict[str, Counter[str]] = defaultdict(Counter)
    document_totals: Counter[str] = Counter()
    for record in records:
        if not record.analyzable:
            continue
        document_totals[record.dataset_id] += 1
        stems, record_surfaces = _document_terms(record, stopwords)
        document_frequency[record.dataset_id].update(stems)
        for stem, counter in record_surfaces.items():
            surfaces[stem].update(counter)

    minimum_df = int(vocabulary_config.get("minimum_document_frequency", 3))
    limit = int(vocabulary_config.get("terms_per_dataset", 40))
    dataset_occurrence = Counter(
        stem
        for dataset in datasets
        for stem, count in document_frequency[dataset.id].items()
        if count >= minimum_df
    )
    dataset_count = len(datasets)

    vocabulary: list[dict[str, Any]] = []
    for dataset in datasets:
        total = document_totals[dataset.id]
        scored: list[tuple[float, str, int]] = []
        for stem, count in document_frequency[dataset.id].items():
            if count < minimum_df:
                continue
            idf = math.log((1 + dataset_count) / (1 + dataset_occurrence[stem])) + 1
            score = (count / max(total, 1)) * idf
            scored.append((score, stem, count))

        top_frequency = sorted(
            ((count / max(total, 1), stem, count) for stem, count in document_frequency[dataset.id].items() if count >= minimum_df),
            reverse=True,
        )[:limit]
        distinctive = sorted(scored, reverse=True)[:limit]
        for kind, values in (("Частотна", top_frequency), ("Характерна", distinctive)):
            for rank, (score, stem, count) in enumerate(values, start=1):
                vocabulary.append({
                    "dataset_id": dataset.id,
                    "dataset": dataset.label,
                    "type": kind,
                    "rank": rank,
                    "term": surfaces[stem].most_common(1)[0][0] if surfaces[stem] else stem,
                    "stem": stem,
                    "titles": count,
                    "percent_of_titles": round(_safe_percent(count, total), 2),
                    "score": round(score, 6),
                })

    common_limit = int(vocabulary_config.get("common_terms", 50))
    common_stems = set.intersection(
        *(
            {stem for stem, count in document_frequency[dataset.id].items() if count >= minimum_df}
            for dataset in datasets
        )
    ) if datasets else set()
    common_ranked = sorted(
        (
            (
                sum(document_frequency[dataset.id][stem] / max(document_totals[dataset.id], 1) for dataset in datasets) / len(datasets),
                stem,
            )
            for stem in common_stems
        ),
        reverse=True,
    )[:common_limit]
    common: list[dict[str, Any]] = []
    for rank, (average_share, stem) in enumerate(common_ranked, start=1):
        for dataset in datasets:
            count = document_frequency[dataset.id][stem]
            common.append({
                "rank": rank,
                "term": surfaces[stem].most_common(1)[0][0] if surfaces[stem] else stem,
                "stem": stem,
                "average_share": round(average_share * 100, 2),
                "dataset_id": dataset.id,
                "dataset": dataset.label,
                "titles": count,
                "percent_of_titles": round(_safe_percent(count, document_totals[dataset.id]), 2),
            })
    return vocabulary, common


def similarity_rows(
    datasets: list[Dataset],
    records: list[Record],
    similarity_config: dict[str, Any],
) -> list[dict[str, Any]]:
    if not similarity_config.get("enabled", True):
        return []
    active = [record for record in records if record.analyzable]
    if len(active) < 2:
        return []

    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.neighbors import NearestNeighbors

    vectorizer = TfidfVectorizer(
        analyzer="char_wb",
        ngram_range=(
            int(similarity_config.get("character_ngram_min", 3)),
            int(similarity_config.get("character_ngram_max", 5)),
        ),
        lowercase=True,
        min_df=2,
        max_features=int(similarity_config.get("max_features", 40000)),
        sublinear_tf=True,
    )
    matrix = vectorizer.fit_transform(record.title for record in active)
    indices_by_dataset: dict[str, list[int]] = defaultdict(list)
    for index, record in enumerate(active):
        indices_by_dataset[record.dataset_id].append(index)

    minimum_score = float(similarity_config.get("minimum_score", 0.28))
    per_case = int(similarity_config.get("matches_per_case", 3))
    results: list[dict[str, Any]] = []
    for left_dataset, right_dataset in combinations(datasets, 2):
        left_indices = indices_by_dataset[left_dataset.id]
        right_indices = indices_by_dataset[right_dataset.id]
        if not left_indices or not right_indices:
            continue
        if len(left_indices) <= len(right_indices):
            query_indices, target_indices = left_indices, right_indices
        else:
            query_indices, target_indices = right_indices, left_indices
        neighbors = min(max(per_case * 4, 10), len(target_indices))
        model = NearestNeighbors(n_neighbors=neighbors, metric="cosine", algorithm="brute", n_jobs=-1)
        model.fit(matrix[target_indices])
        distances, positions = model.kneighbors(matrix[query_indices])
        for query_offset, query_index in enumerate(query_indices):
            accepted = 0
            for distance, target_position in zip(distances[query_offset], positions[query_offset]):
                score = 1.0 - float(distance)
                if score < minimum_score:
                    continue
                left_record = active[query_index]
                right_record = active[target_indices[int(target_position)]]
                results.append({
                    "similarity": round(score, 4),
                    "dataset_a": left_record.dataset_label,
                    "reference_a": left_record.description_reference,
                    "sheet_a": left_record.sheet_name,
                    "case_a": left_record.case_id,
                    "title_a": left_record.title,
                    "dataset_b": right_record.dataset_label,
                    "reference_b": right_record.description_reference,
                    "sheet_b": right_record.sheet_name,
                    "case_b": right_record.case_id,
                    "title_b": right_record.title,
                })
                accepted += 1
                if accepted >= per_case:
                    break
    results.sort(key=lambda row: (-row["similarity"], row["dataset_a"], row["case_a"]))
    return results[:2000]


def build_analysis(
    base_dir: Path,
    datasets: list[Dataset],
    descriptions: list[Description],
    records: list[Record],
    categories: list[Category],
    config: dict[str, Any],
) -> dict[str, Any]:
    dataset_by_id = {dataset.id: dataset for dataset in datasets}
    vocabulary, common = vocabulary_tables(
        datasets,
        records,
        base_dir / "config",
        config.get("vocabulary", {}),
    )
    topic_unclassified = topic_unclassified_rows(records)
    return {
        "overview": overview_rows(datasets, descriptions, records),
        "classification_coverage": classification_coverage_rows(datasets, records),
        "topic_unclassified": topic_unclassified,
        "unclassified": [
            row for row in topic_unclassified
            if row["classification_status"] == "Не класифіковано"
        ],
        "descriptions": description_rows(descriptions, records, dataset_by_id),
        "chronology": chronology_rows(datasets, records, config.get("chronology", {})),
        "categories": category_rows(datasets, records, categories),
        "vocabulary": vocabulary,
        "common_vocabulary": common,
        "similarities": similarity_rows(datasets, records, config.get("similarity", {})),
    }
