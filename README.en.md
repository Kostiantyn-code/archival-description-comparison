# Archival Description Comparison

[![Tests](https://github.com/Kostiantyn-code/archival-description-comparison/actions/workflows/tests.yml/badge.svg)](https://github.com/Kostiantyn-code/archival-description-comparison/actions/workflows/tests.yml)
![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

[Українська версія](README.md)

A local Python tool for comparative chronological, thematic, and lexical
analysis of archival finding aids stored in Excel workbooks. It produces XLSX
and CSV tables, PNG/SVG charts, a JSON run manifest, and a self-contained HTML
report without sending research data to external services.

Development version: `0.3.0.dev0`; latest release: `v0.2.0`.

## Data model

| Level | Interpretation |
|---|---|
| Excel row | An archival file or a service row |
| Worksheet | A separate archival inventory |
| Workbook | A logical dataset compared with other workbooks |

Worksheets remain distinguishable in tables and description-level charts.
Workbook summaries, comparative charts, and vocabulary aggregate all sheets.
One dataset may contain different archives and fonds, including a reconstructed
documentary complex. Archive/fond/inventory metadata belong to each worksheet.
Record IDs use `dataset:sheet:row-N` to distinguish repeated case numbers; they
change if source rows are reordered.

## Features

- any number of `.xlsx` workbooks, with a minimum of two;
- multi-sheet workbook aggregation;
- chronology and dataset-size comparisons, plus a separate series per worksheet;
- stacked bars split by description for dataset size, classification, and categories;
- category panels for compared workbooks sit side by side on a common percentage
  scale, with descriptions stacked inside each bar and one grouped legend entry
  per workbook; each bar uses that workbook's analyzable titles as its denominator;
- description chronology overlays all inventories of a workbook in one panel;
  each yearly value is a percentage of all analyzable cases in that particular
  inventory, on a shared 0–100% scale;
- dictionary-based Ukrainian and Russian title classification;
- separate coverage and unclassified-title reports;
- frequent, distinctive, and shared vocabulary;
- cross-dataset title similarity based on character TF-IDF;
- document-form mentions: counts, percentages and mentions per 1,000 titles;
- category and macroblock co-classification: counts, rates per 1,000 and lift,
  with an offline network selector and cross-dataset table for each selected pair;
- XLSX, CSV, HTML, PNG, SVG, and JSON output;
- entirely local processing.

## Requirements and quick start

Python 3.10 or newer is required.

On Windows, place at least two workbooks in `input` and double-click
`run.bat`. On its first launch it installs dependencies into
`%LOCALAPPDATA%\archival-description-comparison\venv`. Later launches reuse
that environment, including from another extracted ZIP; changes to
`requirements.txt` trigger an update. An existing `.venv` beside `run.bat`
takes precedence. On Linux or macOS:

```bash
chmod +x run.sh
./run.sh
```

Manual setup:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python main.py
```

Results are written to a timestamped directory under `output`.

## Expected worksheet structure

Optional metadata may precede the table:

| A | B |
|---|---|
| Архів / Archive | Archive name |
| Фонд / Fonds | Fonds number |
| Опис / Inventory | Inventory number |
| Мова / Language | `uk` or `ru` |

The data table should contain the following columns:

```text
File number | File title | Inclusive dates | Number of pages | Notes
```

The loader recognizes the Ukrainian and Russian archival column labels used by
the source workbooks. It scans the first 25 rows for the table header and can
recover several common shifted-row patterns.

## Configuration

Dataset names, references, and date boundaries can be defined in
[`config/datasets.yaml`](config/datasets.yaml). Analysis, similarity, and
report settings are located in [`config/analysis.yaml`](config/analysis.yaml).

Each thematic percentage uses the number of analyzable titles in that dataset
as its denominator. Because classification is multi-label, category shares may
sum to more than 100%. See [`docs/methodology.md`](docs/methodology.md) for the
methodological notes.

The `*_by_description.csv` files and matching sheets in `comparison.xlsx`
provide annual counts, classification states, and thematic categories per
worksheet. Category tables give both the worksheet and workbook denominators.
The main bar charts show absolute counts with each bar stacked from its descriptions;
the aggregate chronology and workbook panels with overlaid inventories remain separate plots.

## Visualization language

Ukrainian remains the default. For English illustrations, change
[`config/analysis.yaml`](config/analysis.yaml):

```yaml
visualization:
  language: en             # uk (default) or en
  bilingual_labels: true   # false: translation only; true: English (original)
  label_overrides:
    datasets:
      dataset_a: "Reconstructed documentary collection"
    dataset_short:
      dataset_a: "Collection A"
    archives:
      "Назва архіву": "Archive name"
    descriptions:
      "dataset_a:Опис 1": "Archive name, fond 230: inventory 1"
```

This localizes titles, axes, legends and category labels in all six static
charts, their HTML captions and links, and the theme network's interface,
nodes and scope selector. In English mode, `bilingual_labels` appends original
thematic labels and proper names in parentheses. Interface captions and
structural reference terms use only the selected language. Analysis,
dictionaries, CSV/XLSX, the main HTML tables, case titles and source references
in network examples retain their original data.

Researcher-supplied dataset and archive names require explicit translations.
`datasets` and `dataset_short` use IDs from `datasets.yaml` or `overview.csv`;
short labels fall back to the `datasets` translation when no separate short
translation is provided. `archives` uses exact source archive names.
`descriptions` overrides an inventory reference in individual legends and the
network selector, using `dataset_id:sheet_name`. The category chart's grouped
legend uses archive/fond/inventory metadata. Arbitrary sheet names remain
unchanged; generated names `Опис N` become `Inventory N` in the network.
Unknown labels retain their original text.

The presentation catalog [`config/visualization.en.json`](config/visualization.en.json)
is separate from analytical dictionaries. Overrides also support `categories`,
`category_short`, `macroblocks`, `macroblock_short` and `document_types` by
stable ID. The run manifest records language settings and presentation source
hashes. Each chart is saved as PNG and SVG with text labels; SVG supports
scaling and editing for journal illustrations. Both formats follow
`reports.charts`; the HTML report displays each chart once.

## Document types and theme links

New outputs are `document_types.csv`, `document_types_by_description.csv`,
`document_type_mentions.csv`, `theme_links.csv`, `theme_links_by_description.csv`
and `theme_link_cases.csv`, with six corresponding XLSX sheets. A percentage
heatmap is saved as `figures/document_types.png`. The standalone
`figures/theme_links.html` is linked from the main report and follows
`reports.html` independently of `reports.charts`.

Each type counts once per analyzable title. Mentions describe title wording,
not verified documents inside a case. Subject categories only form pairs;
context-only assignments are excluded. All datasets use the same taxonomy and
include zero-count pairs. For N analyzable titles, marginals A and B and shared
count S, rate = S/N × 1000 and lift = S×N/(A×B). A zero denominator yields a
blank, not zero. Rare-pair lift is descriptive, not evidence of significance or
causality. The network uses fixed node positions and shared metric scales;
its lift view filters counts below max(3, ceil(N×0.0015)), while all pairs remain
in CSV/XLSX. The manifest records `document_type_rules_version`.

## Testing

See the [refactoring audit](docs/refactoring-audit.md) (Ukrainian) for
output equivalence checks and the summary-table benchmark.

```bash
python -m compileall -q main.py src tests
python -m unittest discover -s tests -v
```

## Limitations

Dictionary coverage determines classification coverage. Similarity scores
identify lexically close titles and do not prove that archival files are
identical. The program analyzes finding-aid titles, not the full text of the
documents.

## Related project

The classification dictionaries and rules originate from
[`archival-description-analysis`](https://github.com/Kostiantyn-code/archival-description-analysis).

## Citation, contributions, and license

Citation metadata is available in [`CITATION.cff`](CITATION.cff). Contribution
guidelines are in [`CONTRIBUTING.md`](CONTRIBUTING.md). The source code is
released under the [MIT License](LICENSE). Research workbooks and generated
outputs are not included in the repository.
