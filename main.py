from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from src import __version__
from src.analysis import build_analysis
from src.classification import classify_records, load_language_dictionaries
from src.loader import load_all, safe_load_yaml
from src.reports import create_charts, write_csv_reports, write_html_report, write_manifest, write_workbook


SCRIPT_VERSION = __version__


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Порівняльний контент-аналіз кількох архівних описів у форматі XLSX."
    )
    parser.add_argument("--input-dir", help="Каталог із *.xlsx (типово: input)")
    parser.add_argument("--output-dir", help="Кореневий каталог результатів (типово: output)")
    parser.add_argument("--run-name", help="Фіксована назва папки запуску")
    parser.add_argument("--no-similarity", action="store_true", help="Не шукати подібні справи")
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
    try:
        import yaml
    except ImportError as error:
        raise SystemExit("Не встановлено PyYAML. Виконайте: python -m pip install -r requirements.txt") from error

    config = safe_load_yaml(BASE_DIR / "config" / "analysis.yaml", yaml)
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

    print(f"=== ПОРІВНЯННЯ АРХІВНИХ ОПИСІВ — {SCRIPT_VERSION} ===")
    print(f"Вхідний каталог: {input_dir}")
    print(f"Результати:      {run_dir}")

    print("\n1. Завантаження тематичних словників...")
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
    print(f"   Категорій: {len(categories)}; версія словників: {dictionary_version}")

    print("2. Читання книг і об'єднання аркушів у логічні масиви...")
    datasets, descriptions, records, issues = load_all(
        input_dir,
        datasets_config,
        config.get("chronology", {}),
    )
    for dataset in datasets:
        count = sum(record.analyzable for record in records if record.dataset_id == dataset.id)
        description_count = sum(item.dataset_id == dataset.id for item in descriptions)
        print(f"   {dataset.short_label}: {count:,} заголовків; {description_count} опис(и)".replace(",", " "))

    print("3. Тематична класифікація...")
    started = time.perf_counter()
    classify_records(
        records,
        language_categories,
        language_ambiguities,
        list(macroblock_labels),
        progress=_progress,
    )
    print(f"   Час класифікації: {time.perf_counter() - started:.1f} с")

    print("4. Порівняльні таблиці, лексика та подібні справи...")
    analysis = build_analysis(BASE_DIR, datasets, descriptions, records, categories, config)

    output_files: list[Path] = []
    reports = config.get("reports", {})
    if reports.get("csv", True):
        write_csv_reports(tables_dir, analysis, records, issues)
        output_files.extend(sorted(tables_dir.glob("*.csv")))
    if reports.get("xlsx", True):
        workbook_path = run_dir / "comparison.xlsx"
        write_workbook(workbook_path, datasets, descriptions, records, issues, categories, analysis)
        output_files.append(workbook_path)
    chart_files: list[str] = []
    if reports.get("charts", True):
        chart_files = create_charts(figures_dir, datasets, descriptions, categories, analysis)
        output_files.extend(figures_dir / filename for filename in chart_files)
    if reports.get("html", True):
        html_path = run_dir / "report.html"
        write_html_report(html_path, datasets, analysis, chart_files, issues)
        output_files.append(html_path)

    manifest_path = run_dir / "run_manifest.json"
    write_manifest(manifest_path, BASE_DIR, datasets, dictionary_version, output_files)
    output_files.append(manifest_path)
    print("\nГОТОВО")
    print(f"Логічних масивів: {len(datasets)}")
    print(f"Описів-аркушів:   {len(descriptions)}")
    print(f"Справ у реєстрі:  {sum(record.status == 'case' for record in records):,}".replace(",", " "))
    print(f"Подібних пар:     {len(analysis['similarities']):,}".replace(",", " "))
    print(f"Повна папка:      {run_dir}")
    return run_dir


if __name__ == "__main__":
    try:
        main()
    except FileExistsError as error:
        print(f"\nПОМИЛКА: папка запуску вже існує: {error.filename}")
        raise SystemExit(2)
    except Exception as error:
        print(f"\nКРИТИЧНА ПОМИЛКА: {error}")
        raise
