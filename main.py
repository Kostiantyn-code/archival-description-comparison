from __future__ import annotations

import argparse
import time
from datetime import datetime
from pathlib import Path

import yaml


BASE_DIR = Path(__file__).resolve().parent
from src import __version__
from src.localization import Visualization
from src.analysis import build_analysis
from src.classification import classify_records, load_language_dictionaries
from src.comparative_profiles import write_theme_network
from src.loader import load_all, safe_load_yaml
from src.charts import create_charts
from src.reports import write_csv_reports, write_html_report, write_manifest, write_workbook


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare archival finding aids in two or more XLSX workbooks."
    )
    parser.add_argument("--input-dir", help="Directory containing *.xlsx workbooks (default: input)")
    parser.add_argument("--output-dir", help="Results directory (default: output)")
    parser.add_argument("--run-name", help="Fixed name for this run directory")
    parser.add_argument("--no-similarity", action="store_true", help="Skip title-similarity search")
    return parser.parse_args()


def _path(value: str | None, fallback: str) -> Path:
    selected = Path(value or fallback)
    return selected if selected.is_absolute() else BASE_DIR / selected


def _progress(current: int, total: int) -> None:
    percent = current / total * 100 if total else 100
    print(f"\r   {percent:6.2f}% ({current:,}/{total:,})".replace(",", " "), end="", flush=True)
    if current == total:
        print()


def main() -> Path:
    args = parse_args()
    config = safe_load_yaml(BASE_DIR / "config" / "analysis.yaml", yaml)
    visual = Visualization.from_config(config)
    datasets_config = safe_load_yaml(BASE_DIR / "config" / "datasets.yaml", yaml)
    if args.no_similarity:
        config.setdefault("similarity", {})["enabled"] = False

    input_dir = _path(args.input_dir, str(config.get("input_dir", "input")))
    output_root = _path(args.output_dir, str(config.get("output_dir", "output")))
    run_name = args.run_name or datetime.now().strftime("run_%Y%m%d_%H%M%S")
    run_dir = output_root / run_name
    tables_dir = run_dir / "tables"
    figures_dir = run_dir / "figures"
    run_dir.mkdir(parents=True, exist_ok=False)

    print(f"=== ARCHIVAL DESCRIPTION COMPARISON — {__version__} ===")
    print(f"Input directory: {input_dir}")
    print(f"Output:      {run_dir}")

    print("\n1. Loading thematic dictionaries...")
    language_categories = {}
    language_ambiguities = {}
    macroblock_labels = {}
    dictionary_version = ""
    for language in ("uk", "ru"):
        categories, ambiguities, labels, version = load_language_dictionaries(BASE_DIR, yaml, language)
        language_categories[language] = categories
        language_ambiguities[language] = ambiguities
        macroblock_labels = labels
        dictionary_version = version
    categories = language_categories["uk"]
    print(f"   Categories: {len(categories)}; dictionary version: {dictionary_version}")

    print("2. Reading workbooks and their inventories...")
    datasets, descriptions, records, issues = load_all(
        input_dir,
        datasets_config,
        config.get("chronology", {}),
    )
    for dataset in datasets:
        count = sum(record.analyzable for record in records if record.dataset_id == dataset.id)
        description_count = sum(item.dataset_id == dataset.id for item in descriptions)
        print(f"   {dataset.short_label}: {count:,} titles; {description_count} inventories".replace(",", " "))

    print("3. Classifying case titles...")
    started = time.perf_counter()
    classify_records(
        records,
        language_categories,
        language_ambiguities,
        list(macroblock_labels),
        progress=_progress,
    )
    print(f"   Classification time: {time.perf_counter() - started:.1f} s")

    print("4. Building comparative tables, vocabulary and title similarities...")
    analysis = build_analysis(BASE_DIR, datasets, descriptions, records, categories, config, macroblock_labels)

    output_files: list[Path] = []
    reports = config.get("reports", {})
    if reports.get("csv", True):
        write_csv_reports(tables_dir, analysis, records, issues)
        output_files.extend(sorted(tables_dir.glob("*.csv")))
    if reports.get("xlsx", True):
        workbook_path = run_dir / "comparison.xlsx"
        write_workbook(workbook_path, datasets, descriptions, records, issues, categories, analysis, visual)
        output_files.append(workbook_path)
    chart_files: list[str] = []
    if reports.get("charts", True):
        chart_files = create_charts(figures_dir, datasets, descriptions, categories, analysis, visual)
        output_files.extend(figures_dir / filename for filename in chart_files)
        output_files.extend(sorted(figures_dir.glob("*.svg")))
    if reports.get("html", True):
        network_path = figures_dir / "theme_links.html"
        write_theme_network(network_path, analysis["theme_network"], visual, datasets, descriptions)
        output_files.append(network_path)
        html_path = run_dir / "report.html"
        write_html_report(html_path, datasets, analysis, chart_files, issues, visual)
        output_files.append(html_path)

    manifest_path = run_dir / "run_manifest.json"
    write_manifest(manifest_path, BASE_DIR, datasets, dictionary_version, output_files, visual)
    output_files.append(manifest_path)
    print("\nCOMPLETE")
    print(f"Datasets: {len(datasets)}")
    print(f"Inventories:   {len(descriptions)}")
    print(f"Case records:  {sum(record.status == 'case' for record in records):,}".replace(",", " "))
    print(f"Similar pairs:     {len(analysis['similarities']):,}".replace(",", " "))
    print(f"Results:      {run_dir}")
    return run_dir


if __name__ == "__main__":
    try:
        main()
    except FileExistsError as error:
        print(f"\nERROR: run directory already exists: {error.filename}")
        raise SystemExit(2)
    except Exception as error:
        print(f"\nERROR: {error}")
        raise
