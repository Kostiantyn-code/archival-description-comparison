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

import re
from dataclasses import dataclass
from .models import Category, Dataset, Description, Record
from .localization import Visualization


DOCUMENT_TYPE_RULES_VERSION = "0.1-draft"


@dataclass(frozen=True)
class DocumentType:
    id: str
    label: str
    pattern: re.Pattern[str]


_RULES = (
    ("correspondence", "Correspondence and letters",
     r"\b(?:листуван[а-яіїєґ]*|лист(?:и|а|ів|ами|ах|ом|у|і)?|переписк[а-яё]*|письм(?:о|а|у|ом|е|ами|ах)|писем)\b"),
    ("reports", "Official reports and dispatches",
     r"\b(?:рапорт[а-яіїєґё]*|донесенн[а-яіїєґ]*|донесени[а-яё]*)\b"),
    ("instructions", "Instructions and directives",
     r"\b(?:припис(?:и|ів|а|у|ом|ами|ах)?|розпоряджен[а-яіїєґ]*|предписани[а-яё]*|распоряжени[а-яё]*)\b"),
    ("circulars", "Circulars",
     r"\b(?:обіжник[а-яіїєґ]*|циркуляр[а-яіїєґё]*)\b"),
    ("petitions", "Petitions and requests",
     r"\b(?:клопотанн[а-яіїєґ]*|проханн[а-яіїєґ]*|прошенн[а-яіїєґ]*|ходатайств[а-яё]*|прошени[а-яё]*)\b"),
    ("statements", "Statements and reports",
     r"\b(?:відомост[а-яіїєґ]*|звіт[а-яіїєґ]*|ведомост[а-яё]*|отч[её]т[а-яё]*)\b"),
    ("decrees", "Decrees and orders",
     r"\b(?:указ[а-яіїєґё]*|наказ[а-яіїєґ]*|приказ[а-яё]*)\b"),
    ("memoranda", "Reports and memoranda",
     r"\b(?:доповід[а-яіїєґ]*|записк[а-яіїєґё]*|доклад[а-яё]*)\b"),
    ("lists", "Lists and registers",
     r"\b(?:списк[а-яіїєґё]*|список|реєстр[а-яіїєґ]*|реестр[а-яё]*)\b"),
    ("minutes", "Minutes and journals",
     r"\b(?:протокол[а-яіїєґё]*|журнал[а-яіїєґё]*)\b"),
    ("certificates", "Passports and certificates",
     r"\b(?:паспорт[а-яіїєґё]*|свідоцтв[а-яіїєґ]*|свидетельств[а-яё]*)\b"),
    ("plans", "Plans, projects and cost estimates",
     r"\b(?:план(?:и|ів|у|ом|ах|ами|ов|а|е)?|проєкт[а-яіїєґ]*|проект[а-яё]*|кошторис[а-яіїєґ]*|смет[а-яё]*)\b"),
    ("complaints", "Complaints and applications",
     r"\b(?:скарг[а-яіїєґ]*|заяв(?:а|и|ою|і|ами|ах)|жалоб[а-яё]*|заявлени[а-яё]*)\b"),
    ("acts", "Acts and resolutions",
     r"\b(?:акт(?:и|ів|а|ом|у|ами|ах|ов|ы)?|постанов[а-яіїєґ]*|постановлени[а-яё]*)\b"),
    ("notifications", "Notifications and telegrams",
     r"\b(?:повідомленн[а-яіїєґ]*|телеграм[а-яіїєґё]*|уведомлени[а-яё]*|сообщени[а-яё]*)\b"),
)

DOCUMENT_TYPES = tuple(
    DocumentType(id, label, re.compile(pattern, re.IGNORECASE))
    for id, label, pattern in _RULES
)


def match_document_types(title: str) -> list[tuple[DocumentType, str]]:
    """Return each mentioned type once, with an example matched expression."""
    matches = []
    for document_type in DOCUMENT_TYPES:
        match = document_type.pattern.search(title)
        if match:
            matches.append((document_type, match.group(0)))
    return matches


SCOPE_COLUMNS = [
    ("dataset_id", 'Dataset ID'), ("dataset", 'Dataset'),
    ("file_name", 'File'), ("sheet_name", 'Worksheet'),
    ("archive", 'Archive'), ("fond", 'Fond'), ("inventory", 'Inventory'),
]
RECORD_COLUMNS = SCOPE_COLUMNS + [
    ("record_uid", 'Record ID'), ("excel_row", 'Excel row'),
    ("case_id", 'Case number'), ("title", 'Title'),
]
TYPE_COLUMNS = [
    ("document_type_id", 'Type ID'), ("document_type", 'Document type'),
    ("titles_total", 'All titles'),
    ("titles_with_type", 'Titles with mention'),
    ("percent_of_titles", 'Share of case titles, %'),
    ("per_1000_titles", 'Per 1,000 case titles'),
]
PAIR_COLUMNS = [
    ("level", 'Level'), ("theme_a_id", 'Theme A ID'), ("theme_a", 'Theme A'),
    ("theme_b_id", 'Theme B ID'), ("theme_b", 'Theme B'),
]
LINK_COLUMNS = PAIR_COLUMNS + [
    ("titles_total", 'All titles'),
    ("theme_a_titles", 'With theme A'), ("theme_b_titles", 'With theme B'),
    ("shared_titles", 'Shared titles'),
    ("shared_per_1000_titles", 'Shared per 1,000'),
    ("relative_frequency", 'Lift'),
]
REPORT_COLUMNS = {
    "document_types": SCOPE_COLUMNS[:3] + TYPE_COLUMNS,
    "document_types_by_description": SCOPE_COLUMNS + TYPE_COLUMNS,
    "document_type_mentions": RECORD_COLUMNS + TYPE_COLUMNS[:2] + [("matched_text", 'Matched expression')],
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
                            "ref": f"{record.description_reference}, case {record.case_id}; "
                                   f"{record.sheet_name}, row {record.excel_row}"})
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
        "__SCOPE__", html.escape(visual.text('Comparison of datasets')))
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(page.replace("__DATA__", payload), encoding="utf-8")
