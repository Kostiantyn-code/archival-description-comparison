# Archival Description Comparison

[![Tests](https://github.com/Kostiantyn-code/archival-description-comparison/actions/workflows/tests.yml/badge.svg)](https://github.com/Kostiantyn-code/archival-description-comparison/actions/workflows/tests.yml)
[Українською](#українською) · Python 3.10+ · [MIT](LICENSE)

Compare archival finding aids locally: chronology, thematic coverage, vocabulary,
lexically similar titles, document-form mentions and theme co-classification.
**One workbook = one dataset; one worksheet = one archival inventory.** A dataset
can combine inventories from different archives and fonds (archival record groups).
Source titles and archive names are preserved.

## Start

Download and extract the project, put **at least two `.xlsx` workbooks** in `input`,
then run `run.bat` on Windows or `sh run.sh` on Linux/macOS. Python must be installed;
the launcher installs dependencies on first use and when requirements change.
Manual installation:

```bash
python -m venv .venv
# Activate .venv, then:
python -m pip install -r requirements.txt
python main.py
```

Open `output/run_*/report.html`. Processing and the interactive report work offline
after installation; research data are not sent to external services.

## Input

Each worksheet has five columns **in this order**:

| Case number | Case title | Inclusive dates | Leaf count | Notes |
|---|---|---|---|---|
| 1 | Листування про відкриття школи | 1850–1852 | 12 | |

Optional two-column metadata before the header: `Archive`, `Fond` (or `Fonds`),
`Inventory`, `Language`. Ukrainian and Russian equivalents are also recognized.
Use **`uk` or `ru`** for the language of the entire worksheet; otherwise a heuristic
uses the first 100 titles and defaults to `uk` on a tie. Header detection searches
the first 25 rows. See [methodology](docs/methodology.md) for recovery rules and limits.

**Analysis languages: Ukrainian and Russian.** English is the report language,
not an English classifier. Croatian/Italian analysis is not implemented here.

## Settings and results

- [analysis.yaml](config/analysis.yaml): date limits, vocabulary, similarity and outputs.
  Reports default to English; set `visualization.language: uk` for Ukrainian.
  `bilingual_labels: true` adds Ukrainian taxonomy labels to English figures.
- [datasets.yaml](config/datasets.yaml): optional dataset IDs, labels, references and date limits.
  Translate proper names explicitly with `visualization.label_overrides` if needed.
- [categories.yaml](config/categories.yaml): shared category IDs and English labels.
  [uk.yaml](dictionaries/uk.yaml) and [ru.yaml](dictionaries/ru.yaml) contain the existing
  weighted rules, ambiguity masks and stopwords, one file per language.

| Output | Use |
|---|---|
| `report.html` | Overview, figures and selected tables |
| `comparison.xlsx` | All tables, case register and validation log |
| `tables/*.csv` | Machine-readable tables with stable English fields and generated labels |
| `figures/*.png`, `*.svg` | Figures for inspection and journal illustrations |
| `figures/theme_links.html` | Interactive network, counts, rates and title examples |
| `run_manifest.json` | Input/source hashes, versions, settings and output list |

Workbook and HTML captions follow the selected report language. Archival wording
is never automatically translated. Use `--input-dir`, `--output-dir`, `--run-name`
and `--no-similarity` to override a run; see `python main.py --help`.

Category shares use each dataset's analyzable titles; multiple labels can sum to
more than 100%. Similarity is lexical, does not translate between languages and
does not prove that files are identical. Document-form mentions describe titles,
not the verified contents of a file. Lift is descriptive, not a significance test.

## Development

```bash
python -m compileall -q main.py src tests
python -m unittest discover -s tests -v
```

[Methodology and dictionary format](docs/methodology.md) ·
[Contributing](CONTRIBUTING.md) · [Citation](CITATION.cff) ·
[Related analysis tool](https://github.com/Kostiantyn-code/archival-description-analysis)

## Українською

Локальне порівняння архівних описів: хронологія, тематичне покриття, лексика,
подібні заголовки, згадки типів документів і співкласифікація тем.
**Одна книга Excel = один масив; один аркуш = один архівний опис.** Масив може
об'єднувати описи з різних архівів і фондів. Заголовки й назви архівів зберігаються.

### Запуск

Завантажте й розпакуйте проєкт, покладіть **щонайменше дві книги `.xlsx`** у `input`.
У Windows запустіть `run.bat`, у Linux/macOS — `sh run.sh`. Потрібен установлений
Python 3.10+. Залежності встановлюються під час першого запуску й після зміни
`requirements.txt`. Для ручного встановлення скористайтеся командами вище.
Відкрийте `output/run_*/report.html`. Після встановлення аналіз і звіт працюють
офлайн; дослідницькі дані не передаються зовнішнім сервісам.

### Вхідні дані

П'ять колонок **у фіксованому порядку**: номер справи, заголовок, крайні дати,
кількість аркушів, примітки. Перед таблицею можна додати двоколонкові метадані
`Архів`, `Фонд`, `Опис`, `Мова`; англійські й російські відповідники також підтримуються.
Мова **всього аркуша — `uk` або `ru`**. Без явного значення використовується
евристика за першими 100 заголовками; за рівності голосів обирається `uk`.
Шапка таблиці шукається в перших 25 рядках.

**Мови аналізу: українська й російська.** Англійська є мовою звіту.
Аналіз хорватською та італійською в цій програмі ще не реалізовано.

### Налаштування й результати

Параметри аналізу містяться в `config/analysis.yaml`, назви й межі масивів —
у `config/datasets.yaml`. Типова мова звітів — англійська; для української
задайте `visualization.language: uk`. `bilingual_labels: true` додає українські
назви категорій до англійських ілюстрацій. Переклади власних назв задаються явно
через `visualization.label_overrides`.

Спільні ідентифікатори категорій визначає `config/categories.yaml`. Правила,
маски неоднозначності й стоп-слова об'єднано у `dictionaries/uk.yaml` і `ru.yaml`.

`report.html` містить огляд та ілюстрації; `comparison.xlsx` — усі таблиці,
реєстр справ і журнал перевірки; `tables/*.csv` — машинозчитувані дані зі сталими
англійськими полями й програмними мітками. PNG/SVG містяться в `figures`,
інтерактивна мережа — у `figures/theme_links.html`. `run_manifest.json` фіксує
контрольні суми, версії, налаштування та перелік результатів. Підписи Excel і HTML
відповідають обраній мові; архівні тексти автоматично не перекладаються.

Для окремого запуску доступні `--input-dir`, `--output-dir`, `--run-name`,
`--no-similarity`; довідка — `python main.py --help`.

Частки категорій обчислюються від заголовків відповідного масиву; за множинної
класифікації їх сума може перевищувати 100%. Лексична подібність не перекладає
між мовами й не доводить тотожності справ. Згадки типів документів стосуються
слів у заголовках, а не перевіреного складу справ. Lift є описовим показником.
Методику, обмеження, формат словників і перевірки описано в
[методичних примітках англійською](docs/methodology.md). Команди перевірки наведено вище.
Ліцензія — MIT; метадані цитування — `CITATION.cff`.
