# Archival Description Comparison

[![Tests](https://github.com/Kostiantyn-code/archival-description-comparison/actions/workflows/tests.yml/badge.svg)](https://github.com/Kostiantyn-code/archival-description-comparison/actions/workflows/tests.yml)
![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

[Українська версія](README.md)

A local Python tool for comparative chronological, thematic, and lexical
analysis of archival finding aids stored in Excel workbooks. It produces XLSX
and CSV tables, PNG charts, a JSON run manifest, and a self-contained HTML
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

## Features

- any number of `.xlsx` workbooks, with a minimum of two;
- multi-sheet workbook aggregation;
- chronology and dataset-size comparisons, plus a separate series per worksheet;
- stacked bars split by description for dataset size, classification, and categories;
- category bars for compared workbooks sit side by side, with descriptions
  stacked inside each bar and separate color families for fonds in one workbook;
  each bar uses that workbook's analyzable titles as its percentage denominator;
- description chronology overlays all inventories of a workbook in one panel;
- dictionary-based Ukrainian and Russian title classification;
- separate coverage and unclassified-title reports;
- frequent, distinctive, and shared vocabulary;
- cross-dataset title similarity based on character TF-IDF;
- XLSX, CSV, HTML, PNG, and JSON output;
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

## Testing

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
