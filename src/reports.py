"""CSV, Excel, HTML та графічні результати порівняння."""
from __future__ import annotations

import csv
import hashlib
import html
import json
import platform
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from . import __version__
from .models import Category, Dataset, Description, Issue, Record


NAVY = "1F4E78"
BLUE = "D9EAF7"
LIGHT_BLUE = "EAF3F8"
PALE = "F6F8FA"
WHITE = "FFFFFF"
TEXT = "202020"
GRID = "D9E1E8"
WARNING = "FFF2CC"
ERROR = "F4CCCC"


def _write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    if not rows:
        path.write_text("", encoding="utf-8-sig")
        return
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(rows[0]), delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def record_rows(records: list[Record]) -> list[dict[str, Any]]:
    return [{
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
        "start_year": record.start_year,
        "end_year": record.end_year,
        "pages": record.pages,
        "notes": record.notes,
        "status": record.status,
        "language": record.language,
        "categories": " | ".join(record.category_labels),
        "category_ids": " | ".join(record.categories),
        "context_categories": " | ".join(record.context_category_labels),
        "review_flags": " | ".join(record.review_flags),
    } for record in records]


def issue_rows(issues: list[Issue]) -> list[dict[str, Any]]:
    return [{
        "level": issue.level,
        "file_name": issue.file_name,
        "sheet_name": issue.sheet_name,
        "row": issue.row,
        "field": issue.field,
        "value": issue.value,
        "message": issue.message,
    } for issue in issues]


def write_csv_reports(
    tables_dir: Path,
    analysis: dict[str, Any],
    records: list[Record],
    issues: list[Issue],
) -> None:
    tables_dir.mkdir(parents=True, exist_ok=True)
    mapping = {
        "overview.csv": analysis["overview"],
        "classification_coverage.csv": analysis["classification_coverage"],
        "topic_unclassified_cases.csv": analysis["topic_unclassified"],
        "unclassified_cases.csv": analysis["unclassified"],
        "descriptions.csv": analysis["descriptions"],
        "chronology.csv": analysis["chronology"],
        "chronology_by_description.csv": analysis["chronology_by_description"],
        "categories.csv": analysis["categories"],
        "categories_by_description.csv": analysis["categories_by_description"],
        "classification_by_description.csv": analysis["classification_by_description"],
        "vocabulary.csv": analysis["vocabulary"],
        "common_vocabulary.csv": analysis["common_vocabulary"],
        "similarities.csv": analysis["similarities"],
        "records.csv": record_rows(records),
        "issues.csv": issue_rows(issues),
    }
    for filename, rows in mapping.items():
        _write_csv(tables_dir / filename, rows)


def _title(sheet, text: str, subtitle: str = "") -> int:
    sheet.sheet_view.showGridLines = False
    sheet.cell(2, 1, text)
    sheet.cell(2, 1).font = Font(name="Arial", size=14, bold=True, color=TEXT)
    sheet.cell(3, 1).fill = PatternFill("solid", fgColor=NAVY)
    sheet.row_dimensions[3].height = 3
    row = 5
    if subtitle:
        sheet.cell(4, 1, subtitle)
        sheet.cell(4, 1).font = Font(name="Arial", size=10, italic=True, color="666666")
        row = 6
    return row


def _tabular_sheet(
    workbook: Workbook,
    title: str,
    heading: str,
    columns: list[tuple[str, str]],
    rows: list[dict[str, Any]],
    subtitle: str = "",
    percent_keys: set[str] | None = None,
    decimal_keys: set[str] | None = None,
) -> None:
    sheet = workbook.create_sheet(title)
    start_row = _title(sheet, heading, subtitle)
    percent_keys = percent_keys or set()
    decimal_keys = decimal_keys or set()
    for column_index, (_, label) in enumerate(columns, start=1):
        sheet.cell(3, column_index).fill = PatternFill("solid", fgColor=NAVY)
        cell = sheet.cell(start_row, column_index, label)
        cell.font = Font(name="Arial", size=10, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_index, row in enumerate(rows, start=start_row + 1):
        for column_index, (key, _) in enumerate(columns, start=1):
            value = row.get(key)
            cell = sheet.cell(row_index, column_index, value)
            cell.font = Font(name="Arial", size=10, color=TEXT)
            cell.alignment = Alignment(vertical="top", wrap_text=key in {"title", "title_a", "title_b", "message", "reference"})
            if key in percent_keys and isinstance(value, (int, float)):
                cell.value = value / 100
                cell.number_format = "0.00%"
            elif key in decimal_keys and isinstance(value, (int, float)):
                cell.number_format = "0.0000"
        if row_index % 2 == 0:
            for column_index in range(1, len(columns) + 1):
                sheet.cell(row_index, column_index).fill = PatternFill("solid", fgColor=PALE)
    if rows:
        end_row = start_row + len(rows)
        table = Table(
            displayName="T_" + "".join(character for character in title if character.isalnum())[:20] + str(len(workbook.worksheets)),
            ref=f"A{start_row}:{get_column_letter(len(columns))}{end_row}",
        )
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)
    sheet.freeze_panes = f"A{start_row + 1}"
    sheet.auto_filter.ref = f"A{start_row}:{get_column_letter(len(columns))}{start_row + max(1, len(rows))}"
    sheet.row_dimensions[start_row].height = 32
    for column_index, (key, label) in enumerate(columns, start=1):
        sample = [str(row.get(key, "") or "") for row in rows[:400]]
        longest = max([len(label), *(min(len(value), 80) for value in sample)], default=len(label))
        width = min(max(longest + 2, 10), 55)
        if key in {"title", "title_a", "title_b", "message"}:
            width = 52
        sheet.column_dimensions[get_column_letter(column_index)].width = width


def _chronology_wide(datasets: list[Dataset], rows: list[dict[str, Any]]) -> tuple[list[tuple[str, str]], list[dict[str, Any]]]:
    columns = [("year", "Рік")] + [(dataset.id, dataset.short_label) for dataset in datasets]
    return columns, rows


def _category_wide(
    datasets: list[Dataset],
    categories: list[Category],
    rows: list[dict[str, Any]],
) -> tuple[list[tuple[str, str]], list[dict[str, Any]], set[str]]:
    lookup = {(row["dataset_id"], row["category_id"]): row for row in rows}
    columns = [("category", "Категорія")]
    percent_keys: set[str] = set()
    for dataset in datasets:
        columns.append((f"{dataset.id}_cases", f"{dataset.short_label}: справ"))
        key = f"{dataset.id}_percent"
        columns.append((key, f"{dataset.short_label}: частка"))
        percent_keys.add(key)
    wide_rows: list[dict[str, Any]] = []
    for category in categories:
        item: dict[str, Any] = {"category": category.label}
        for dataset in datasets:
            source = lookup[(dataset.id, category.id)]
            item[f"{dataset.id}_cases"] = source["cases"]
            item[f"{dataset.id}_percent"] = source["percent_of_titles"]
        wide_rows.append(item)
    return columns, wide_rows, percent_keys


def _common_wide(
    datasets: list[Dataset], rows: list[dict[str, Any]]
) -> tuple[list[tuple[str, str]], list[dict[str, Any]], set[str]]:
    grouped: dict[int, dict[str, Any]] = {}
    percent_keys: set[str] = set()
    for row in rows:
        item = grouped.setdefault(row["rank"], {
            "rank": row["rank"],
            "term": row["term"],
            "average_share": row["average_share"],
        })
        item[f"{row['dataset_id']}_titles"] = row["titles"]
        item[f"{row['dataset_id']}_percent"] = row["percent_of_titles"]
    columns = [("rank", "№"), ("term", "Термін"), ("average_share", "Середня частка")]
    percent_keys.add("average_share")
    for dataset in datasets:
        columns.append((f"{dataset.id}_titles", f"{dataset.short_label}: заголовків"))
        key = f"{dataset.id}_percent"
        columns.append((key, f"{dataset.short_label}: частка"))
        percent_keys.add(key)
    return columns, list(grouped.values()), percent_keys


def write_workbook(
    path: Path,
    datasets: list[Dataset],
    descriptions: list[Description],
    records: list[Record],
    issues: list[Issue],
    categories: list[Category],
    analysis: dict[str, Any],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    _tabular_sheet(
        workbook, "Огляд", "Порівняння архівних описів",
        [
            ("dataset", "Логічний масив"), ("reference", "Склад масиву"),
            ("descriptions", "Описів"), ("cases", "Справ"),
            ("analyzable_titles", "Заголовків в аналізі"), ("withdrawn", "Вибулих"),
            ("start_year", "Початковий рік"), ("end_year", "Кінцевий рік"),
            ("pages_total", "Аркушів разом"), ("pages_mean", "Середнє аркушів"),
            ("pages_median", "Медіана аркушів"), ("missing_titles", "Без заголовка"),
            ("missing_dates", "Без дат"), ("missing_pages", "Без аркушів"),
            ("classified_titles", "Тематично класифіковано"),
            ("classification_coverage_percent", "Покриття класифікацією"),
            ("context_only_titles", "Лише контекст"),
            ("context_only_percent", "Лише контекст, %"),
            ("unclassified_titles", "Не класифіковано"),
            ("unclassified_percent", "Не класифіковано, %"),
        ],
        analysis["overview"],
        subtitle="Кожен файл Excel трактовано як один логічний масив; його аркуші об'єднано.",
        percent_keys={
            "classification_coverage_percent", "context_only_percent",
            "unclassified_percent",
        },
    )
    _tabular_sheet(
        workbook, "Класифікація", "Покриття тематичною класифікацією",
        [
            ("dataset", "Логічний масив"), ("reference", "Склад масиву"),
            ("analyzable_titles", "Усього справ (100%)"),
            ("subject_classified", "Тематично класифіковано"),
            ("subject_classified_percent", "Тематично класифіковано, %"),
            ("context_only", "Лише контекст"),
            ("context_only_percent", "Лише контекст, %"),
            ("unclassified", "Не класифіковано"),
            ("unclassified_percent", "Не класифіковано, %"),
        ],
        analysis["classification_coverage"],
        subtitle=(
            "За 100% взято всі аналізовані справи кожного логічного масиву "
            "окремо; три стани взаємовиключні й разом дають 100%."
        ),
        percent_keys={
            "subject_classified_percent", "context_only_percent",
            "unclassified_percent",
        },
    )
    _tabular_sheet(
        workbook, "Некласифіковані", "Справи без предметної категорії",
        [
            ("dataset", "Логічний масив"), ("file_name", "Файл"),
            ("sheet_name", "Аркуш"), ("excel_row", "Рядок Excel"),
            ("archive", "Архів"), ("fond", "Фонд"),
            ("inventory", "Опис"), ("case_id", "№ справи"),
            ("title", "Заголовок"), ("dates_raw", "Крайні дати"),
            ("classification_status", "Статус"),
            ("context_categories", "Контекстні категорії"),
            ("review_flags", "Позначки для перевірки"),
        ],
        analysis["topic_unclassified"],
        subtitle=(
            "Усі справи без предметної категорії. Фільтр у колонці «Статус» "
            "відокремлює записи «Лише контекст» від повністю некласифікованих."
        ),
    )
    _tabular_sheet(
        workbook, "Описи", "Контрольні підсумки за аркушами-описами",
        [
            ("dataset", "Логічний масив"), ("file_name", "Файл"), ("sheet_name", "Аркуш"),
            ("archive", "Архів"), ("fond", "Фонд"), ("inventory", "Опис"),
            ("cases", "Справ"), ("analyzable_titles", "Заголовків в аналізі"),
            ("withdrawn", "Вибулих"), ("classified_titles", "Класифіковано"),
            ("missing_dates", "Без дат"), ("missing_pages", "Без аркушів"),
        ], analysis["descriptions"],
    )
    chronology_columns, chronology_data = _chronology_wide(datasets, analysis["chronology"])
    _tabular_sheet(
        workbook, "Хронологія", "Присутність справ за роками",
        chronology_columns, chronology_data,
        subtitle="Справу з діапазоном дат зараховано до кожного року в межах діапазону.",
    )
    _tabular_sheet(
        workbook, "Хронологія описів", "Присутність справ кожного опису за роками",
        [("dataset", "Масив"), ("sheet_name", "Аркуш"), ("reference", "Опис"),
         ("year", "Рік"), ("cases", "Справ")],
        analysis["chronology_by_description"],
        subtitle="Діапазон дат зараховано до кожного року; описи однієї книги залишаються окремими.",
    )
    category_columns, category_data, category_percent = _category_wide(datasets, categories, analysis["categories"])
    _tabular_sheet(
        workbook, "Категорії", "Тематична структура масивів",
        category_columns, category_data,
        subtitle=(
            "Для кожного масиву 100% — усі його аналізовані справи. Один заголовок "
            "може належати до кількох категорій, тому сума часток може перевищувати 100%."
        ),
        percent_keys=category_percent,
    )
    _tabular_sheet(
        workbook, "Категорії описів", "Категорії за окремими описами",
        [("dataset", "Масив"), ("sheet_name", "Аркуш"), ("reference", "Опис"),
         ("category", "Категорія"), ("cases", "Справ"),
         ("description_titles_total", "Справ в описі"),
         ("percent_of_description_titles", "Частка опису"),
         ("dataset_titles_total", "Справ у масиві"),
         ("percent_of_dataset_titles", "Частка масиву")],
        analysis["categories_by_description"],
        subtitle="Для частки масиву за 100% взято всі його справи; для частки опису — справи відповідного аркуша.",
        percent_keys={"percent_of_description_titles", "percent_of_dataset_titles"},
    )
    _tabular_sheet(
        workbook, "Класифікація описів", "Покриття класифікацією за описами",
        [("dataset", "Масив"), ("sheet_name", "Аркуш"), ("reference", "Опис"),
         ("analyzable_titles", "Справ в описі"),
         ("subject_classified", "Тематично класифіковано"),
         ("context_only", "Лише контекст"), ("unclassified", "Не класифіковано"),
         ("subject_classified_percent", "Класифіковано, %"),
         ("context_only_percent", "Лише контекст, %"),
         ("unclassified_percent", "Не класифіковано, %")],
        analysis["classification_by_description"],
        subtitle="Три взаємовиключні стани; частки в таблиці обчислено від справ відповідного опису.",
        percent_keys={"subject_classified_percent", "context_only_percent", "unclassified_percent"},
    )
    _tabular_sheet(
        workbook, "Характерна лексика", "Частотна та характерна лексика",
        [
            ("dataset", "Логічний масив"), ("type", "Список"), ("rank", "Місце"),
            ("term", "Термін"), ("titles", "Заголовків"),
            ("percent_of_titles", "Частка заголовків"), ("score", "Оцінка характерності"),
        ], analysis["vocabulary"],
        percent_keys={"percent_of_titles"}, decimal_keys={"score"},
    )
    common_columns, common_data, common_percent = _common_wide(datasets, analysis["common_vocabulary"])
    _tabular_sheet(
        workbook, "Спільна лексика", "Спільна лексика всіх масивів",
        common_columns, common_data, percent_keys=common_percent,
    )
    _tabular_sheet(
        workbook, "Подібні справи", "Лексично подібні справи між масивами",
        [
            ("similarity", "Подібність"),
            ("dataset_a", "Масив A"), ("reference_a", "Посилання A"),
            ("case_a", "Справа A"), ("title_a", "Заголовок A"),
            ("dataset_b", "Масив B"), ("reference_b", "Посилання B"),
            ("case_b", "Справа B"), ("title_b", "Заголовок B"),
        ], analysis["similarities"], decimal_keys={"similarity"},
    )
    _tabular_sheet(
        workbook, "Справи", "Нормалізований реєстр справ",
        [
            ("dataset", "Логічний масив"), ("sheet_name", "Аркуш"), ("excel_row", "Рядок Excel"),
            ("archive", "Архів"), ("fond", "Фонд"), ("inventory", "Опис"),
            ("case_id", "№ справи"), ("title", "Заголовок"), ("dates_raw", "Крайні дати"),
            ("start_year", "Початковий рік"), ("end_year", "Кінцевий рік"),
            ("pages", "Аркушів"), ("status", "Статус"), ("language", "Мова"),
            ("categories", "Категорії"),
        ], record_rows(records),
    )
    _tabular_sheet(
        workbook, "Перевірка", "Журнал перевірки вхідних даних",
        [
            ("level", "Рівень"), ("file_name", "Файл"), ("sheet_name", "Аркуш"),
            ("row", "Рядок"), ("field", "Поле"), ("value", "Значення"),
            ("message", "Повідомлення"),
        ], issue_rows(issues),
    )
    check_sheet = workbook["Перевірка"]
    for row in range(6, check_sheet.max_row + 1):
        level = check_sheet.cell(row, 1).value
        fill = ERROR if level == "ERROR" else WARNING if level == "WARNING" else None
        if fill:
            for column in range(1, check_sheet.max_column + 1):
                check_sheet.cell(row, column).fill = PatternFill("solid", fgColor=fill)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def _description_colors(descriptions: list[Description]) -> dict[tuple[str, str], str]:
    """Give each archive/fond in each workbook its own family of shades."""
    import matplotlib.colors as mcolors

    families = [
        ("#366A9A", "#D8E8F5"), ("#92591D", "#F4E0BE"),
        ("#557A46", "#D6E5D1"), ("#70477C", "#E3D5E7"),
        ("#A23B72", "#F1D4E1"), ("#1D7777", "#D0E9E7"),
    ]
    grouped: dict[tuple[str, str, str], list[Description]] = defaultdict(list)
    for item in descriptions:
        grouped[(item.dataset_id, item.archive.strip(), item.fond.strip())].append(item)
    colors: dict[tuple[str, str], str] = {}
    for index, group in enumerate(grouped.values()):
        dark, light = (mcolors.to_rgb(value) for value in families[index % len(families)])
        for shade, item in enumerate(group):
            factor = 0.38 if len(group) == 1 else shade / max(len(group) - 1, 1)
            colors[(item.dataset_id, item.sheet_name)] = mcolors.to_hex(
                tuple(start * (1 - factor) + end * factor for start, end in zip(dark, light))
            )
    return colors


def _description_label(item: Description) -> str:
    if item.fond:
        prefix = f"{item.archive}, " if item.archive else ""
        detail = f"опис {item.inventory}" if item.inventory else item.sheet_name
        return f"{prefix}фонд {item.fond}: {detail}"
    return item.reference


def _fond_summary(descriptions: list[Description]) -> str:
    """Name all fonds and their inventory numbers within one workbook."""
    groups: dict[tuple[str, str], list[str]] = defaultdict(list)
    for item in descriptions:
        groups[(item.archive, item.fond)].append(item.inventory or item.sheet_name)
    return "; ".join(
        f"{archive + ': ' if archive else ''}Фонд {fond or 'без номера'}: "
        f"{'Опис' if len(numbers) == 1 else 'Описи'} {', '.join(dict.fromkeys(numbers))}"
        for (archive, fond), numbers in groups.items()
    )


def _stacked_description_bars(
    ax, labels: list[str], descriptions: list[Description],
    values: dict[tuple[str, str], list[int]], colors: dict[tuple[str, str], str],
    total_label: str = "Усього", *, horizontal: bool = True,
) -> None:
    from matplotlib.colors import to_rgb

    totals = [
        sum(values[(item.dataset_id, item.sheet_name)][index] for item in descriptions)
        for index in range(len(labels))
    ]
    offsets = [0] * len(labels)
    maximum = max(totals, default=0)
    for item in descriptions:
        key = (item.dataset_id, item.sheet_name)
        segment = values[key]
        color = colors[key]
        style = {"color": color, "label": _description_label(item), "edgecolor": "white", "linewidth": 0.25}
        bars = (ax.barh(labels, segment, left=offsets, **style) if horizontal
                else ax.bar(labels, segment, bottom=offsets, **style))
        visible = [
            f"{value:,}".replace(",", " ")
            if value >= 10 and value / max(total, 1) >= 0.025
            and value / max(maximum, 1) >= (0.018 if horizontal else 0.035) else ""
            for value, total in zip(segment, totals)
        ]
        red, green, blue = to_rgb(color)
        label_color = "white" if 0.2126 * red + 0.7152 * green + 0.0722 * blue < 0.55 else "#202020"
        ax.bar_label(bars, labels=visible, label_type="center", fontsize=7, color=label_color)
        offsets = [left + value for left, value in zip(offsets, segment)]
    for index, total in enumerate(totals):
        if total:
            content = f"{total_label}: {total:,}".replace(",", " ")
            if horizontal:
                ax.text(total + max(maximum * 0.012, 0.5), index,
                        content, va="center", fontsize=8)
            else:
                ax.text(index, total + max(maximum * 0.012, 0.5),
                        content, ha="center", fontsize=8)
    if horizontal:
        ax.set_xlim(0, maximum * 1.24 if maximum else 1)
        ax.invert_yaxis()
    else:
        ax.set_ylim(0, maximum * 1.15 if maximum else 1)
    ax.grid(axis="x" if horizontal else "y", alpha=0.22)
    ax.set_axisbelow(True)
    ax.legend(title="Архівний опис", frameon=False, fontsize=8,
              loc="lower center", bbox_to_anchor=(0.5, 1.02),
              ncol=min(len(descriptions), 3))


def _category_dataset_label(dataset: Dataset, descriptions: list[Description]) -> str:
    """One legend entry per workbook, listing all its fonds and inventories."""
    groups: dict[tuple[str, str], list[str]] = defaultdict(list)
    for item in descriptions:
        groups[(item.archive.strip(), item.fond.strip())].append(item.inventory or item.sheet_name)
    return "\n".join(
        f"{archive + ', ' if archive else ''}ф. {fond}, оп. {', '.join(dict.fromkeys(numbers))}"
        if fond else f"{archive or dataset.short_label}, оп. {', '.join(dict.fromkeys(numbers))}"
        for (archive, fond), numbers in groups.items()
    ) or dataset.short_label


def _category_panel(
    ax, descriptions: list[Description],
    categories: list[Category], counts: dict[tuple[str, str, str], int],
    colors: dict[tuple[str, str], str], dataset_total: int, maximum: float,
) -> None:
    """Stack inventory shares in one bar per category for one workbook."""
    from matplotlib.colors import to_rgb

    offsets = [0.0] * len(categories)
    case_totals = [0] * len(categories)
    for item in descriptions:
        key = (item.dataset_id, item.sheet_name)
        case_values = [counts[(item.dataset_id, item.sheet_name, category.id)] for category in categories]
        values = [count / max(dataset_total, 1) * 100 for count in case_values]
        bars = ax.barh(range(len(categories)), values, left=offsets,
                       height=0.72, color=colors[key],
                       edgecolor="white", linewidth=0.25)
        red, green, blue = to_rgb(colors[key])
        label_color = "white" if 0.2126 * red + 0.7152 * green + 0.0722 * blue < 0.55 else "#202020"
        segment_labels = []
        for index, (value, cases) in enumerate(zip(values, case_values)):
            segment_labels.append(
                f"{value:.1f}%" if value >= 2.5 and value / max(maximum, 1) >= 0.06 else ""
            )
            offsets[index] += value
            case_totals[index] += cases
        ax.bar_label(bars, labels=segment_labels, label_type="center", fontsize=7, color=label_color)

    for index, (total, cases) in enumerate(zip(offsets, case_totals)):
        if cases:
            formatted_cases = f"{cases:,}".replace(",", " ")
            ax.text(total + max(maximum * 0.014, 0.15), index,
                    f"{total:.1f}% ({formatted_cases})", va="center", fontsize=8)
    ax.set_yticks(range(len(categories)), [category.label for category in categories])
    ax.set_xlim(0, maximum * 1.28 if maximum else 1)
    ax.grid(axis="x", alpha=0.22)
    ax.set_axisbelow(True)


def _category_figure(
    datasets: list[Dataset], descriptions: list[Description],
    categories: list[Category], counts: dict[tuple[str, str, str], int],
    colors: dict[tuple[str, str], str], dataset_totals: dict[str, int],
):
    """Adjacent workbook panels share a percentage scale and one grouped legend."""
    import matplotlib.pyplot as plt
    from matplotlib.legend_handler import HandlerTuple
    from matplotlib.patches import Patch

    columns = min(len(datasets), 2)
    rows = (len(datasets) + columns - 1) // columns
    fig, axes = plt.subplots(rows, columns, sharex=True, sharey=True, squeeze=False,
                             figsize=(max(11, 8.4 * columns), max(6.8, 5.4 * rows)))
    maximum = max((
        sum(counts[(item.dataset_id, item.sheet_name, category.id)]
            for item in descriptions if item.dataset_id == dataset.id)
        / max(dataset_totals[dataset.id], 1) * 100
        for dataset in datasets for category in categories
    ), default=0)
    handles = []
    labels = []
    for index, dataset in enumerate(datasets):
        ax = axes.flat[index]
        subset = [item for item in descriptions if item.dataset_id == dataset.id]
        _category_panel(ax, subset, categories, counts, colors,
                        dataset_totals[dataset.id], maximum)
        ax.set_title(dataset.short_label, loc="left", fontsize=11)
        if index % columns:
            ax.tick_params(axis="y", left=False, labelleft=False)
        handles.append(tuple(Patch(facecolor=colors[(item.dataset_id, item.sheet_name)])
                             for item in subset) or (Patch(facecolor="#999999"),))
        labels.append(_category_dataset_label(dataset, subset))
    for ax in axes.flat[len(datasets):]:
        ax.set_visible(False)
    axes[0][0].invert_yaxis()
    fig.suptitle("Тематичні категорії за описами", fontsize=14, y=0.99)
    fig.legend(handles, labels, handler_map={tuple: HandlerTuple(ndivide=None, pad=0.05)},
               loc="upper center", bbox_to_anchor=(0.55, 0.96), ncol=columns,
               frameon=False, fontsize=9, handlelength=2.2)
    fig.supxlabel("Частка справ відповідного файлу з категорією, %")
    fig.subplots_adjust(left=0.23 if columns == 2 else 0.35,
                        right=0.98, top=0.79, bottom=0.11,
                        wspace=0.12, hspace=0.33)
    return fig


def _description_chronology_panel(
    ax, dataset: Dataset, descriptions: list[Description], years: list[int],
    yearly: dict[tuple[str, str, int], int], colors: dict[tuple[str, str], str],
) -> None:
    subset = [item for item in descriptions if item.dataset_id == dataset.id]
    series = [
        [yearly[(item.dataset_id, item.sheet_name, year)] for year in years]
        for item in subset
    ]
    for item, values in zip(subset, series):
        ax.plot(years, values,
                color=colors[(item.dataset_id, item.sheet_name)], linewidth=1.7,
                label=_description_label(item))
    ax.set_title(dataset.short_label, loc="left", fontsize=10)
    ax.set_ylabel("Справ")
    ax.grid(alpha=0.22)
    ax.legend(title=_fond_summary(subset), fontsize=8, title_fontsize=8,
              frameon=False, loc="lower center", bbox_to_anchor=(0.5, 1.02),
              ncol=min(len(subset), 3))
    if len(series) > 1 and years:
        peaks = [max(values, default=0) for values in series]
        dominant = max(range(len(peaks)), key=peaks.__getitem__)
        smaller_peak = max((peak for index, peak in enumerate(peaks) if index != dominant), default=0)
        if smaller_peak and peaks[dominant] > 5 * smaller_peak:
            from matplotlib.ticker import MaxNLocator

            zoom = ax.inset_axes((0.035, 0.54, 0.37, 0.32))
            active_years = []
            for index, (item, values) in enumerate(zip(subset, series)):
                if index != dominant:
                    zoom.plot(years, values,
                              color=colors[(item.dataset_id, item.sheet_name)], linewidth=1.5)
                    active_years.extend(year for year, count in zip(years, values) if count)
            zoom.set_ylim(0, smaller_peak * 1.13)
            start, end = min(active_years), max(active_years)
            margin = max(1, (end - start) * 0.06)
            zoom.set_xlim(start - margin, end + margin)
            zoom.set_title("Інші описи — збільшена шкала", fontsize=8)
            zoom.xaxis.set_major_locator(MaxNLocator(nbins=4, integer=True))
            zoom.tick_params(labelsize=7)
            zoom.grid(alpha=0.18)
            zoom.set_facecolor("white")


def create_charts(
    figures_dir: Path,
    datasets: list[Dataset],
    descriptions: list[Description],
    categories: list[Category],
    analysis: dict[str, Any],
) -> list[str]:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    figures_dir.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.family": "DejaVu Sans", "axes.titlesize": 13, "axes.labelsize": 10})
    palette = ["#2F6690", "#D17A22", "#548C2F", "#805D93", "#A23B72", "#2A9D8F"]
    created: list[str] = []
    colors = _description_colors(descriptions)

    overview = analysis["overview"]
    labels = [row["short_label"] for row in overview]
    if descriptions:
        description_sizes = {
            (row["dataset_id"], row["sheet_name"]): row["analyzable_titles"]
            for row in analysis["descriptions"]
        }
        fig, ax = plt.subplots(figsize=(12.5, max(4.5, len(datasets) * 0.9 + 2)))
        _stacked_description_bars(
            ax, labels, descriptions,
            {(item.dataset_id, item.sheet_name): [
                description_sizes[(item.dataset_id, item.sheet_name)] if item.dataset_id == dataset.id else 0
                for dataset in datasets
            ] for item in descriptions}, colors,
        )
        ax.set_title("Обсяг порівнюваних масивів за описами", pad=75)
        ax.set_xlabel("Кількість заголовків справ")
        fig.tight_layout()
        filename = "dataset_sizes.png"
        fig.savefig(figures_dir / filename, dpi=180, bbox_inches="tight")
        plt.close(fig)
        created.append(filename)

    chronology = analysis["chronology"]
    if chronology:
        fig, axes = plt.subplots(
            len(datasets), 1,
            figsize=(11, max(4.8, 2.7 * len(datasets))),
            sharex=True,
            squeeze=False,
        )
        years = [row["year"] for row in chronology]
        for index, dataset in enumerate(datasets):
            ax = axes[index][0]
            ax.plot(
                years,
                [row[dataset.id] for row in chronology],
                color=palette[index % len(palette)],
                linewidth=1.8,
            )
            ax.set_title(dataset.short_label, loc="left", fontsize=10)
            ax.set_ylabel("Справ")
            ax.grid(alpha=0.22)
        axes[-1][0].set_xlabel("Рік")
        fig.suptitle("Хронологічний розподіл справ", fontsize=13)
        fig.tight_layout()
        filename = "chronology.png"
        fig.savefig(figures_dir / filename, dpi=180, bbox_inches="tight")
        plt.close(fig)
        created.append(filename)

    chronology_parts = analysis["chronology_by_description"]
    if chronology_parts:
        fig, axes = plt.subplots(
            len(datasets), 1, figsize=(12, max(5.2, 3.7 * len(datasets))),
            sharex=True, squeeze=False,
        )
        yearly = {(row["dataset_id"], row["sheet_name"], row["year"]): row["cases"]
                  for row in chronology_parts}
        years = sorted({row["year"] for row in chronology_parts})
        for index, dataset in enumerate(datasets):
            _description_chronology_panel(
                axes[index][0], dataset, descriptions, years, yearly, colors
            )
        axes[-1][0].set_xlabel("Рік")
        fig.suptitle("Хронологія описів у кожному порівнюваному масиві", fontsize=13)
        fig.tight_layout()
        filename = "chronology_by_description.png"
        fig.savefig(figures_dir / filename, dpi=180, bbox_inches="tight")
        plt.close(fig)
        created.append(filename)

    by_description = {
        (row["dataset_id"], row["sheet_name"]): row
        for row in analysis["classification_by_description"]
    }
    if by_description:
        fig, axes = plt.subplots(len(datasets), 1,
                                 figsize=(11.5, max(5.5, 5.2 * len(datasets))), squeeze=False)
        states = [("subject_classified", "Тематично класифіковано"),
                  ("context_only", "Лише контекст"),
                  ("unclassified", "Не класифіковано")]
        for index, dataset in enumerate(datasets):
            subset = [item for item in descriptions if item.dataset_id == dataset.id]
            ax = axes[index][0]
            _stacked_description_bars(
                ax, [label for _, label in states], subset,
                {(item.dataset_id, item.sheet_name): [
                    by_description[(item.dataset_id, item.sheet_name)][key] for key, _ in states
                ] for item in subset}, colors, horizontal=False,
            )
            ax.set_title(dataset.short_label, loc="left")
            ax.set_ylabel("Кількість справ")
        fig.suptitle("Покриття класифікацією за окремими описами", fontsize=13)
        fig.tight_layout()
        filename = "classification_coverage.png"
        fig.savefig(figures_dir / filename, dpi=180, bbox_inches="tight")
        plt.close(fig)
        created.append(filename)

    category_parts = {
        (row["dataset_id"], row["sheet_name"], row["category_id"]): row["cases"]
        for row in analysis["categories_by_description"]
    }
    if category_parts:
        dataset_totals = {row["dataset_id"]: row["analyzable_titles"] for row in analysis["overview"]}
        fig = _category_figure(datasets, descriptions, categories, category_parts,
                               colors, dataset_totals)
        filename = "categories.png"
        fig.savefig(figures_dir / filename, dpi=180, bbox_inches="tight")
        plt.close(fig)
        created.append(filename)
    return created


def _html_table(columns: list[tuple[str, str]], rows: Iterable[dict[str, Any]], limit: int | None = None) -> str:
    selected = list(rows)
    if limit is not None:
        selected = selected[:limit]
    parts = ["<div class='table-wrap'><table><thead><tr>"]
    parts.extend(f"<th>{html.escape(label)}</th>" for _, label in columns)
    parts.append("</tr></thead><tbody>")
    for row in selected:
        parts.append("<tr>")
        for key, _ in columns:
            value = row.get(key, "")
            if isinstance(value, float):
                text = f"{value:.2f}"
            else:
                text = str(value if value is not None else "")
            parts.append(f"<td>{html.escape(text)}</td>")
        parts.append("</tr>")
    parts.append("</tbody></table></div>")
    return "".join(parts)


def write_html_report(
    path: Path,
    datasets: list[Dataset],
    analysis: dict[str, Any],
    chart_files: list[str],
    issues: list[Issue],
) -> None:
    overview = analysis["overview"]
    issue_counts = Counter(issue.level for issue in issues)
    cards = "".join(
        f"<article class='card'><h3>{html.escape(row['short_label'])}</h3>"
        f"<strong>{row['analyzable_titles']:,}</strong><span>заголовків</span>"
        f"<p>{html.escape(row['reference'])}</p></article>".replace(",", " ")
        for row in overview
    )
    chart_titles = {
        "dataset_sizes.png": "Обсяг масивів: склад кожного стовпчика за описами",
        "chronology.png": "Хронологія масивів",
        "chronology_by_description.png": "Одна панель на книгу: хронологічні ряди описів накладено",
        "classification_coverage.png": "Стан класифікації: стовпчики зі справ окремих описів",
        "categories.png": "Тематичні категорії: сусідні панелі фондів і спільна легенда",
    }
    charts = "".join(
        f"<figure><img src='figures/{html.escape(filename)}' alt='{html.escape(chart_titles.get(filename, filename))}'>"
        f"<figcaption>{html.escape(chart_titles.get(filename, filename))}</figcaption></figure>"
        for filename in chart_files
    )
    top_vocabulary = [row for row in analysis["vocabulary"] if row["type"] == "Характерна" and row["rank"] <= 15]
    content = f"""<!doctype html>
<html lang="uk"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Порівняння архівних описів</title>
<style>
:root{{--navy:#1f4e78;--blue:#d9eaf7;--ink:#202020;--muted:#667085;--line:#d9e1e8}}
*{{box-sizing:border-box}} body{{margin:0;font-family:Arial,sans-serif;color:var(--ink);background:#f5f7fa}}
main{{max-width:1280px;margin:auto;padding:32px}} h1{{margin:0 0 8px;font-size:30px}} h2{{margin-top:42px;border-bottom:3px solid var(--navy);padding-bottom:8px}}
.lead{{color:var(--muted);max-width:900px}} .cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:16px;margin:24px 0}}
.card{{background:white;border:1px solid var(--line);border-radius:10px;padding:18px}} .card h3{{margin:0 0 12px;font-size:17px}} .card strong{{font-size:30px;color:var(--navy)}} .card span{{margin-left:8px;color:var(--muted)}} .card p{{font-size:13px;color:var(--muted)}}
figure{{margin:20px 0;background:white;border:1px solid var(--line);padding:14px;border-radius:10px}} img{{max-width:100%;height:auto;display:block;margin:auto}} figcaption{{font-size:12px;color:var(--muted);text-align:center}}
.table-wrap{{overflow:auto;background:white;border:1px solid var(--line);border-radius:8px}} table{{border-collapse:collapse;width:100%;font-size:14px}} th{{background:var(--navy);color:white;text-align:left;position:sticky;top:0}} th,td{{padding:9px 10px;border-bottom:1px solid var(--line);vertical-align:top}} tr:nth-child(even) td{{background:#f8fafc}}
.note{{background:#fff8e1;border-left:4px solid #d17a22;padding:12px 16px}} footer{{margin-top:48px;color:var(--muted);font-size:13px}}
</style></head><body><main>
<h1>Порівняння архівних описів</h1>
<p class="lead">Одиниця порівняння — логічний масив, представлений одним файлом Excel. На стовпчикових діаграмах кожен стовпчик складено із сегментів відповідних аркушів-описів. Річна хронологія подана як за книгами, так і окремим рядом для кожного опису.</p>
<section class="cards">{cards}</section>
<p class="note">Журнал перевірки: помилок — {issue_counts['ERROR']}, попереджень — {issue_counts['WARNING']}, інформаційних повідомлень — {issue_counts['INFO']}.</p>
<h2>Основні показники</h2>
{_html_table([('dataset','Масив'),('reference','Склад'),('descriptions','Описів'),('analyzable_titles','Заголовків'),('start_year','Від'),('end_year','До'),('classification_coverage_percent','Класифіковано, %')], overview)}
<h2>Покриття тематичною класифікацією</h2>
<p class="lead">За 100% взято всі аналізовані справи кожного логічного масиву окремо.</p>
{_html_table([('dataset','Масив'),('analyzable_titles','Усього справ (100%)'),('subject_classified','Тематично класифіковано'),('subject_classified_percent','Класифіковано, %'),('context_only','Лише контекст'),('context_only_percent','Лише контекст, %'),('unclassified','Не класифіковано'),('unclassified_percent','Не класифіковано, %')], analysis['classification_coverage'])}
<h2>Склад масивів за описами</h2>
{_html_table([('dataset','Масив'),('reference','Архівний опис'),('analyzable_titles','Справ в аналізі'),('classified_titles','Тематично класифіковано'),('withdrawn','Вибулих')], analysis['descriptions'])}
<p class="lead">Діаграма категорій показує частку від справ кожного масиву окремо; її сегменти — внесок описів у цю частку. Інші стовпчикові графіки показують кількості. Таблиці «Категорії описів» і «Класифікація описів» у comparison.xlsx та CSV містять кількості й частки зі знаменниками опису та масиву.</p>
<h2>Візуалізації</h2>{charts}
<h2>Характерна лексика</h2>
{_html_table([('dataset','Масив'),('rank','Місце'),('term','Термін'),('titles','Заголовків'),('percent_of_titles','Частка, %')], top_vocabulary)}
<h2>Найподібніші справи</h2>
{_html_table([('similarity','Подібність'),('reference_a','Посилання A'),('case_a','Справа A'),('title_a','Заголовок A'),('reference_b','Посилання B'),('case_b','Справа B'),('title_b','Заголовок B')], analysis['similarities'], 100)}
<footer>Створено {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}. Повні таблиці містяться у comparison.xlsx та каталозі tables.</footer>
</main></body></html>"""
    path.write_text(content, encoding="utf-8")


def write_manifest(
    path: Path,
    base_dir: Path,
    datasets: list[Dataset],
    dictionary_version: str,
    output_files: list[Path],
) -> None:
    inputs = [dataset.path for dataset in datasets]
    payload = {
        "status": "complete",
        "script_version": __version__,
        "dictionary_version": dictionary_version,
        "completed_at": datetime.now().isoformat(timespec="seconds"),
        "python": platform.python_version(),
        "datasets": [
            {"id": dataset.id, "label": dataset.label, "file": dataset.path.name, "reference": dataset.reference}
            for dataset in datasets
        ],
        "sha256": {
            item.name: hashlib.sha256(item.read_bytes()).hexdigest()
            for item in inputs
        },
        "outputs": [str(item.relative_to(path.parent)) for item in output_files if item.exists()],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
