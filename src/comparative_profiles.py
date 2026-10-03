"""Comparable document-form mentions and subject co-classification profiles.

Counts refer to analyzable titles, never to individual documents in a case.
Every dataset uses the same rules and taxonomy, including zero-count rows.
"""
from __future__ import annotations

import html
import json
from collections import Counter, defaultdict
from copy import deepcopy
from itertools import combinations
from pathlib import Path

from .document_types import DOCUMENT_TYPES, match_document_types
from .models import Category, Dataset, Description, Record
from .visualization import Visualization


SCOPE_COLUMNS = [
    ("dataset_id", "ID масиву"), ("dataset", "Масив"),
    ("file_name", "Файл"), ("sheet_name", "Аркуш"),
    ("archive", "Архів"), ("fond", "Фонд"), ("inventory", "Опис"),
]
RECORD_COLUMNS = SCOPE_COLUMNS + [
    ("record_uid", "ID запису"), ("excel_row", "Рядок Excel"),
    ("case_id", "№ справи"), ("title", "Заголовок"),
]
TYPE_COLUMNS = [
    ("document_type_id", "ID типу"), ("document_type", "Тип документа"),
    ("titles_total", "Усього заголовків"),
    ("titles_with_type", "Заголовків зі згадкою"),
    ("percent_of_titles", "Частка заголовків, %"),
    ("per_1000_titles", "На 1000 заголовків"),
]
PAIR_COLUMNS = [
    ("level", "Рівень"), ("theme_a_id", "ID теми A"), ("theme_a", "Тема A"),
    ("theme_b_id", "ID теми B"), ("theme_b", "Тема B"),
]
LINK_COLUMNS = PAIR_COLUMNS + [
    ("titles_total", "Усього заголовків"),
    ("theme_a_titles", "З темою A"), ("theme_b_titles", "З темою B"),
    ("shared_titles", "Спільних заголовків"),
    ("shared_per_1000_titles", "Спільних на 1000"),
    ("relative_frequency", "Відносна частота (lift)"),
]
REPORT_COLUMNS = {
    "document_types": SCOPE_COLUMNS[:3] + TYPE_COLUMNS,
    "document_types_by_description": SCOPE_COLUMNS + TYPE_COLUMNS,
    "document_type_mentions": RECORD_COLUMNS + TYPE_COLUMNS[:2] + [("matched_text", "Знайдений вираз")],
    "theme_links": SCOPE_COLUMNS[:3] + LINK_COLUMNS,
    "theme_links_by_description": SCOPE_COLUMNS + LINK_COLUMNS,
    "theme_link_cases": RECORD_COLUMNS + PAIR_COLUMNS,
}


def _rate(count: int, total: int, scale: int) -> float | None:
    return round(count / total * scale, 6) if total else None


def _source(record: Record) -> dict:
    return {key: getattr(record, key) for key in (
        "dataset_id", "file_name", "sheet_name", "archive", "fond", "inventory",
        "excel_row", "case_id", "title",
    )} | {"dataset": record.dataset_label, "record_uid": record.uid}


def build_comparative_profiles(
    datasets: list[Dataset], descriptions: list[Description], records: list[Record],
    categories: list[Category], macroblock_labels: dict[str, str] | None = None,
) -> dict:
    result = {key: [] for key in REPORT_COLUMNS}
    category_map = {c.id: c for c in categories}
    blocks = {c.macroblock: (macroblock_labels or {}).get(c.macroblock, c.macroblock)
              for c in categories}
    groups = {"cat": {c.id: c.label for c in categories}, "block": blocks}
    by_dataset = defaultdict(list)
    by_description = defaultdict(list)
    # Key by physical row, so repeated case numbers cannot overwrite evidence.
    types = {}
    assignments = {}
    for record in records:
        if not record.analyzable:
            continue
        by_dataset[record.dataset_id].append(record)
        by_description[(record.dataset_id, record.sheet_name)].append(record)
        matches = match_document_types(record.title)
        types[record.uid] = [kind.id for kind, _ in matches]
        subject = set(record.categories) & category_map.keys()
        assigned = {"cat": subject, "block": {category_map[c].macroblock for c in subject}}
        assignments[record.uid] = assigned
        for kind, matched in matches:
            result["document_type_mentions"].append(_source(record) | {
                "document_type_id": kind.id, "document_type": kind.label, "matched_text": matched,
            })
        for level, ids in assigned.items():
            for a, b in combinations(sorted(ids), 2):
                result["theme_link_cases"].append(_source(record) | {
                    "level": level, "theme_a_id": a, "theme_a": groups[level][a],
                    "theme_b_id": b, "theme_b": groups[level][b],
                })

    scopes = []
    for dataset in datasets:
        meta = {"dataset_id": dataset.id, "dataset": dataset.label, "file_name": dataset.path.name}
        scopes.append((meta, by_dataset[dataset.id], False, dataset.label))
        for description in descriptions:
            if description.dataset_id != dataset.id:
                continue
            source = meta | {key: getattr(description, key)
                             for key in ("sheet_name", "archive", "fond", "inventory")}
            scopes.append((source, by_description[(dataset.id, description.sheet_name)], True,
                           f"{dataset.label} / {description.sheet_name} ({description.reference})"))

    network = {"meta": {
        "cat": groups["cat"], "short": {c.id: c.label[:20] + "…" if len(c.label) > 21 else c.label for c in categories},
        "block": blocks, "blockShort": {key: label[:20] + "…" if len(label) > 21 else label for key, label in blocks.items()},
        "catOrder": list(groups["cat"]), "catBlock": {c.id: c.macroblock for c in categories},
        "scopes": [], "datasets": [],
    }, "scopes": {}}
    for index, (meta, subset, is_description, label) in enumerate(scopes):
        scope_id = f"scope_{index}"
        network["meta"]["scopes"].append([scope_id, label])
        if not is_description:
            network["meta"]["datasets"].append([scope_id, meta["dataset"]])
        network["scopes"][scope_id] = {}
        suffix = "_by_description" if is_description else ""
        total = len(subset)
        type_counts = Counter(kind for record in subset for kind in types[record.uid])
        for kind in DOCUMENT_TYPES:
            count = type_counts[kind.id]
            result["document_types" + suffix].append(meta | {
                "document_type_id": kind.id, "document_type": kind.label,
                "titles_total": total, "titles_with_type": count,
                "percent_of_titles": _rate(count, total, 100),
                "per_1000_titles": _rate(count, total, 1000),
            })
        for level, known in groups.items():
            nodes = Counter()
            pairs = Counter()
            examples = defaultdict(list)
            multi = 0
            for record in subset:
                ids = assignments[record.uid][level]
                nodes.update(ids)
                multi += len(ids) > 1
                for pair in combinations(sorted(ids), 2):
                    pairs[pair] += 1
                    if len(examples[pair]) < 3:
                        examples[pair].append({"title": record.title,
                            "ref": f"{record.description_reference}, спр. {record.case_id}; "
                                   f"{record.sheet_name}, рядок {record.excel_row}"})
            edges = []
            for a, b in combinations(sorted(known), 2):
                shared = pairs[(a, b)]
                denominator = nodes[a] * nodes[b]
                lift = round(shared * total / denominator, 6) if denominator else None
                rate = _rate(shared, total, 1000)
                result["theme_links" + suffix].append(meta | {
                    "level": level, "theme_a_id": a, "theme_a": known[a],
                    "theme_b_id": b, "theme_b": known[b], "titles_total": total,
                    "theme_a_titles": nodes[a], "theme_b_titles": nodes[b],
                    "shared_titles": shared, "shared_per_1000_titles": rate,
                    "relative_frequency": lift,
                })
                edges.append({"a": a, "b": b, "count": shared, "rate": rate,
                              "lift": lift, "examples": examples[(a, b)]})
            network["scopes"][scope_id][level] = {
                "total": total, "multi": multi, "nodes": dict(nodes), "edges": edges,
            }
    result["theme_network"] = network
    return result


def write_theme_network(
    path: Path, network: dict, visual: Visualization | None = None,
    datasets: list[Dataset] | None = None, descriptions: list[Description] | None = None,
) -> None:
    visual = visual or Visualization()
    template = Path(__file__).parent / "templates" / "theme_links.html"
    displayed = network
    if visual.language == "en":
        displayed = deepcopy(network)  # Keep analytical data and original labels immutable.
        meta = displayed["meta"]
        meta["bilingual"] = visual.bilingual
        for field, group in (("cat", "categories"), ("short", "category_short"),
                             ("block", "macroblocks"), ("blockShort", "macroblock_short")):
            meta[field] = {key: visual.label(group, key, label) for key, label in meta[field].items()}
        if datasets is not None and descriptions is not None:
            scope_labels, dataset_labels = [], []
            for dataset in datasets:
                label = visual.dataset(dataset.id, dataset.label)
                scope_labels.append(label)
                dataset_labels.append(label)
                for item in descriptions:
                    if item.dataset_id == dataset.id:
                        key = f"{dataset.id}:{item.sheet_name}"
                        reference = visual.inventory_reference(key, item.archive, item.fond,
                                                               item.inventory, item.reference)
                        scope_labels.append(f"{label} / {visual.inventory_name(item.sheet_name)} ({reference})")
            if len(scope_labels) != len(meta["scopes"]) or len(dataset_labels) != len(meta["datasets"]):
                raise ValueError("Network source metadata does not match the analytical scopes")
            meta["scopes"] = [[key, label] for (key, _), label in zip(meta["scopes"], scope_labels)]
            meta["datasets"] = [[key, label] for (key, _), label in zip(meta["datasets"], dataset_labels)]
    payload = json.dumps(displayed, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    page = visual.template(template.read_text(encoding="utf-8")).replace(
        "__SCOPE__", html.escape(visual.text("Порівняння логічних масивів")))
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(page.replace("__DATA__", payload), encoding="utf-8")
