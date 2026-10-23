"""Comparative figures in PNG and editable SVG."""
from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from textwrap import fill
from typing import Any

from .localization import Visualization
from .models import Category, Dataset, Description
from .comparative_profiles import DOCUMENT_TYPES


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


def _description_label(item: Description, visual: Visualization | None = None) -> str:
    visual = visual or Visualization()
    reference = visual.inventory_reference(f"{item.dataset_id}:{item.sheet_name}",
                                           item.archive, item.fond, item.inventory, item.reference)
    return fill(reference, 55)


def _fond_summary(descriptions: list[Description], visual: Visualization | None = None) -> str:
    visual = visual or Visualization()
    groups: dict[tuple[str, str], list[str]] = defaultdict(list)
    for item in descriptions:
        groups[(item.archive, item.fond)].append(item.inventory or item.sheet_name)
    return fill("; ".join(
        f"{visual.label('archives', archive, archive) + ': ' if archive else ''}"
        f"{visual.text('Fond')} {fond or visual.text('unnumbered')}: "
        f"{visual.text('Inventory' if len(numbers) == 1 else 'Inventories')} "
        f"{', '.join(dict.fromkeys(numbers))}"
        for (archive, fond), numbers in groups.items()), 70)


def _stacked_description_bars(
    ax, labels: list[str], descriptions: list[Description],
    values: dict[tuple[str, str], list[int]], colors: dict[tuple[str, str], str],
    total_label: str = 'Total', *, horizontal: bool = True,
    visual: Visualization | None = None,
) -> None:
    from matplotlib.colors import to_rgb

    visual = visual or Visualization()
    total_label = visual.text(total_label)
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
        style = {"color": color, "label": _description_label(item, visual), "edgecolor": "white", "linewidth": 0.25}
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
    ax.legend(title=visual.text('Archival inventory'), frameon=False, fontsize=8,
              loc="lower center", bbox_to_anchor=(0.5, 1.02),
              ncol=min(len(descriptions), 3))


def _category_dataset_label(dataset: Dataset, descriptions: list[Description],
                            visual: Visualization | None = None) -> str:
    visual = visual or Visualization()
    groups: dict[tuple[str, str], list[str]] = defaultdict(list)
    for item in descriptions:
        groups[(item.archive.strip(), item.fond.strip())].append(item.inventory or item.sheet_name)
    return "\n".join(fill(visual.reference(
        archive or visual.dataset(dataset.id, dataset.short_label, short=True),
        fond, ", ".join(dict.fromkeys(numbers))), 60)
        for (archive, fond), numbers in groups.items()) or visual.dataset(dataset.id, dataset.short_label, short=True)


def _category_panel(
    ax, descriptions: list[Description],
    categories: list[Category], counts: dict[tuple[str, str, str], int],
    colors: dict[tuple[str, str], str], dataset_total: int, maximum: float,
    visual: Visualization | None = None,
) -> None:
    """Stack inventory shares in one bar per category for one workbook."""
    visual = visual or Visualization()
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
    ax.set_yticks(range(len(categories)), [fill(visual.label("categories", category.id, category.label), 44)
                 for category in categories])
    ax.set_xlim(0, maximum * 1.28 if maximum else 1)
    ax.grid(axis="x", alpha=0.22)
    ax.set_axisbelow(True)


def _category_figure(
    datasets: list[Dataset], descriptions: list[Description],
    categories: list[Category], counts: dict[tuple[str, str, str], int],
    colors: dict[tuple[str, str], str], dataset_totals: dict[str, int],
    visual: Visualization | None = None,
):
    """Adjacent workbook panels share a percentage scale and one grouped legend."""
    visual = visual or Visualization()
    import matplotlib.pyplot as plt
    from matplotlib.legend_handler import HandlerTuple
    from matplotlib.patches import Patch

    columns = min(len(datasets), 2)
    rows = (len(datasets) + columns - 1) // columns
    label_lines = sum(fill(visual.label("categories", c.id, c.label), 44).count("\n") + 1
                      for c in categories)
    fig, axes = plt.subplots(rows, columns, sharex=True, sharey=True, squeeze=False,
                             figsize=(max(11, (9.5 if visual.bilingual else 8.4) * columns),
                                      max(6.8, 5.4 * rows) if visual.language == "uk" else
                                      max(6.8, rows * (label_lines * 0.27 + 2.5))))
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
                        dataset_totals[dataset.id], maximum, visual)
        ax.set_title(visual.dataset(dataset.id, dataset.short_label, short=True), loc="left", fontsize=11)
        if index % columns:
            ax.tick_params(axis="y", left=False, labelleft=False)
        handles.append(tuple(Patch(facecolor=colors[(item.dataset_id, item.sheet_name)])
                             for item in subset) or (Patch(facecolor="#999999"),))
        labels.append(_category_dataset_label(dataset, subset, visual))
    for ax in axes.flat[len(datasets):]:
        ax.set_visible(False)
    axes[0][0].invert_yaxis()
    fig.suptitle(visual.text('Thematic categories by inventory'), fontsize=14, y=0.99)
    fig.legend(handles, labels, handler_map={tuple: HandlerTuple(ndivide=None, pad=0.05)},
               loc="upper center", bbox_to_anchor=(0.55, 0.96), ncol=columns,
               frameon=False, fontsize=9, handlelength=2.2)
    fig.supxlabel(visual.text('Share of case files in the corresponding dataset with the category, %'))
    fig.subplots_adjust(left=0.23 if columns == 2 else 0.35,
                        right=0.98, top=0.79, bottom=0.11,
                        wspace=0.12, hspace=0.33)
    return fig


def _description_chronology_panel(
    ax, dataset: Dataset, descriptions: list[Description], years: list[int],
    yearly: dict[tuple[str, str, int], int], colors: dict[tuple[str, str], str],
    description_totals: dict[tuple[str, str], int],
    visual: Visualization | None = None,
) -> None:
    from matplotlib.ticker import PercentFormatter

    visual = visual or Visualization()

    subset = [item for item in descriptions if item.dataset_id == dataset.id]
    for item in subset:
        key = (item.dataset_id, item.sheet_name)
        total = description_totals[key]
        values = [yearly[(item.dataset_id, item.sheet_name, year)] / max(total, 1) * 100
                  for year in years]
        ax.plot(years, values,
                color=colors[key], linewidth=1.7,
                label=_description_label(item, visual))
    ax.set_title(f"{visual.dataset(dataset.id, dataset.short_label, short=True)} — "
                 f"{visual.text('100% for each inventory separately')}",
                 loc="left", fontsize=10)
    ax.set_ylabel(visual.text('Share of case files in the inventory, %'))
    ax.set_ylim(0, 100)
    ax.yaxis.set_major_formatter(PercentFormatter(xmax=100))
    ax.grid(alpha=0.22)
    ax.legend(title=_fond_summary(subset, visual), fontsize=8, title_fontsize=8,
              frameon=False, loc="lower center", bbox_to_anchor=(0.5, 1.13),
              ncol=min(len(subset), 3))


def create_charts(
    figures_dir: Path,
    datasets: list[Dataset],
    descriptions: list[Description],
    categories: list[Category],
    analysis: dict[str, Any],
    visual: Visualization | None = None,
) -> list[str]:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    visual = visual or Visualization()
    figures_dir.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.family": "DejaVu Sans", "axes.titlesize": 13, "axes.labelsize": 10})
    palette = ["#2F6690", "#D17A22", "#548C2F", "#805D93", "#A23B72", "#2A9D8F"]
    created: list[str] = []
    colors = _description_colors(descriptions)

    def save_chart(fig, filename: str) -> None:
        fig.savefig(figures_dir / filename, dpi=180, bbox_inches="tight")
        with plt.rc_context({"svg.fonttype": "none"}):
            fig.savefig(figures_dir / Path(filename).with_suffix(".svg"), bbox_inches="tight")
        plt.close(fig)
        created.append(filename)

    overview = analysis["overview"]
    labels = [visual.dataset(row["dataset_id"], row["short_label"], short=True)
              for row in overview]
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
            ] for item in descriptions}, colors, visual=visual,
        )
        ax.set_title(visual.text('Size of compared datasets by inventory'), pad=75)
        ax.set_xlabel(visual.text('Number of case titles'))
        fig.tight_layout()
        save_chart(fig, "dataset_sizes.png")

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
            ax.set_title(visual.dataset(dataset.id, dataset.short_label, short=True), loc="left", fontsize=10)
            ax.set_ylabel(visual.text('Case files'))
            ax.grid(alpha=0.22)
        axes[-1][0].set_xlabel(visual.text('Year'))
        fig.suptitle(visual.text('Chronological distribution of case files'), fontsize=13)
        fig.tight_layout()
        save_chart(fig, "chronology.png")

    chronology_parts = analysis["chronology_by_description"]
    if chronology_parts:
        description_totals = {
            (row["dataset_id"], row["sheet_name"]): row["analyzable_titles"]
            for row in analysis["descriptions"]
        }
        fig, axes = plt.subplots(
            len(datasets), 1, figsize=(12, max(5.2, 3.7 * len(datasets))),
            sharex=True, squeeze=False,
        )
        yearly = {(row["dataset_id"], row["sheet_name"], row["year"]): row["cases"]
                  for row in chronology_parts}
        years = sorted({row["year"] for row in chronology_parts})
        for index, dataset in enumerate(datasets):
            _description_chronology_panel(
                axes[index][0], dataset, descriptions, years, yearly, colors,
                description_totals, visual,
            )
        axes[-1][0].set_xlabel(visual.text('Year'))
        fig.suptitle(visual.text('Inventory chronology within each compared dataset'), fontsize=13, y=1.04)
        fig.tight_layout(h_pad=2.5)
        save_chart(fig, "chronology_by_description.png")

    by_description = {
        (row["dataset_id"], row["sheet_name"]): row
        for row in analysis["classification_by_description"]
    }
    if by_description:
        fig, axes = plt.subplots(len(datasets), 1,
                                 figsize=(11.5, max(5.5, 5.2 * len(datasets))), squeeze=False)
        states = [("subject_classified", 'Subject classified'),
                  ("context_only", 'Context only'),
                  ("unclassified", 'Unclassified')]
        for index, dataset in enumerate(datasets):
            subset = [item for item in descriptions if item.dataset_id == dataset.id]
            ax = axes[index][0]
            _stacked_description_bars(
                ax, [visual.text(label) for _, label in states], subset,
                {(item.dataset_id, item.sheet_name): [
                    by_description[(item.dataset_id, item.sheet_name)][key] for key, _ in states
                ] for item in subset}, colors, horizontal=False, visual=visual,
            )
            ax.set_title(visual.dataset(dataset.id, dataset.short_label, short=True), loc="left")
            ax.set_ylabel(visual.text('Number of case files'))
        fig.suptitle(visual.text('Classification coverage by inventory'), fontsize=13)
        fig.tight_layout()
        save_chart(fig, "classification_coverage.png")

    category_parts = {
        (row["dataset_id"], row["sheet_name"], row["category_id"]): row["cases"]
        for row in analysis["categories_by_description"]
    }
    if category_parts:
        dataset_totals = {row["dataset_id"]: row["analyzable_titles"] for row in analysis["overview"]}
        fig = _category_figure(datasets, descriptions, categories, category_parts,
                               colors, dataset_totals, visual)
        save_chart(fig, "categories.png")
    if analysis.get("document_types") and datasets:
        import numpy as np
        values = {(row["dataset_id"], row["document_type_id"]): row["percent_of_titles"]
                  for row in analysis["document_types"]}
        matrix = np.array([[values.get((d.id, kind.id))
                            if values.get((d.id, kind.id)) is not None else np.nan
                            for d in datasets] for kind in DOCUMENT_TYPES])
        label_lines = sum(fill(visual.label("document_types", kind.id, kind.label), 44).count("\n") + 1
                          for kind in DOCUMENT_TYPES)
        fig, ax = plt.subplots(figsize=(max(9, len(datasets) * 1.6 + (7 if visual.bilingual else 5)),
                                       9 if visual.language == "uk" else max(9, label_lines * 0.3)))
        cmap = plt.get_cmap("Blues").copy()
        cmap.set_bad("#dddddd")
        heat = ax.imshow(matrix, aspect="auto", vmin=0, vmax=100, cmap=cmap)
        ax.set_xticks(range(len(datasets)), [visual.dataset(d.id, d.short_label, short=True) for d in datasets], rotation=30, ha="right")
        ax.set_yticks(range(len(DOCUMENT_TYPES)), [fill(visual.label("document_types", kind.id, kind.label), 44)
                      for kind in DOCUMENT_TYPES])
        for i in range(len(DOCUMENT_TYPES)):
            for j in range(len(datasets)):
                value = matrix[i, j]
                ax.text(j, i, "—" if np.isnan(value) else f"{value:.1f}", ha="center", va="center",
                        color="white" if value > 55 else "#202020", fontsize=9)
        ax.set_title(visual.text('Document types mentioned, % of dataset titles'), pad=16)
        fig.colorbar(heat, ax=ax, label=visual.text('Share of case titles, %'))
        fig.tight_layout()
        save_chart(fig, "document_types.png")
    return created
