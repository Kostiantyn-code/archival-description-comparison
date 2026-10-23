# Methodology and reproducibility

This tool compares **finding-aid titles**, not the full texts of archival documents.
Its categories and lexical indicators are exploratory and require corpus-specific
manual evaluation before supporting substantive historical claims.

## Units and input recovery

A workbook is one dataset and a worksheet is one inventory. A dataset can combine
inventories belonging to different archives or fonds. Archive, fond and inventory
metadata are read separately from each worksheet. Source rows are identified by
`dataset_id:sheet_name:row-N`; repeated case numbers stay distinct, but reordering
rows changes these IDs.

The loader expects five fixed columns: case number, title, inclusive dates,
leaf count, notes. It recognizes English, Ukrainian and Russian metadata/header
labels and searches the first 25 rows for the header. Metadata are scanned in the
first 15 rows. Spreadsheet values are read with `data_only=True`, so input formulas
require cached values; plain text is preferable. Dates are interpreted as years
1700–2099 within configured bounds, not as full normalized dates. A single numeric
leaf count is accepted; ambiguous counts are left blank and logged.

Blank rows and repeated headers are skipped. Year headings can supply a year to
following cases, and a case number shifted into a year-heading row can be joined
to the next title. Dates or leaf counts beside a year heading can fill missing
fields in the preceding case. Unnumbered rows are generally headings; only the
supported shifted-row pattern creates a synthetic case number. These rules do
not recover every malformed worksheet. Review `issues.csv`, record counts and
source rows before analysis. Withdrawn cases, inventory entries and empty titles
do not contribute to analytical denominators.

## Language and dictionary format

`Language | uk` or `Language | ru` applies to the entire worksheet. Ukrainian and
Russian metadata equivalents are accepted. Without metadata, the first 100 titles
after the header vote: a confidently identified title contributes two votes;
selected diagnostic words in otherwise undetermined titles contribute one.
A tie selects Ukrainian. Mixed-language titles do not switch dictionaries.
Set the language explicitly for short or mixed corpora.

There are 12 shared category IDs and four macroblocks in `config/categories.yaml`.
English labels are canonical; Ukrainian report labels are in
`config/localization.json`. Each `dictionaries/{uk,ru}.yaml` contains:

| Field | Meaning |
|---|---|
| `language` | Analysis language, `uk` or `ru` |
| `categories` | Rules keyed by the shared category `id` |
| `ambiguities` | Cross-category masking and conditional rules |
| `stopwords` | Lexical-analysis exclusions; not classification exclusions |

Within a category, `strong_phrases` contribute 4 points, `strong_terms` 3 and
`contextual_terms` 1. Co-occurrence `context_rules` have an explicit score
(default 3); a satisfied ambiguity condition adds 3. The usual assignment
threshold (`minimum_score`) is 3. `context_only_phrases` and `context_only_terms`
record context without adding subject scores. Contextual terms also appear in
context evidence. In Russian, context-only score accumulation cannot by itself
assign a subject category: at least one non-context item is required.
`exclude_phrases` and ambiguity masks suppress matches before scoring.
`review_terms` flag near-threshold cases. Multiple matching aliases can contribute
separate points; do not deduplicate them without evaluating the resulting scores.
Light stemming is heuristic, not full lemmatization; UK/RU caches are separate.

The `personal_files` category describes documentation type rather than a subject.
It can co-occur with a subject, for example education or policing. Explicit
personnel-file/list section headings can supply categories to name-only titles;
letter/year headings preserve that section, while a new substantive heading ends
it. Inherited evidence is marked `inventory section` and the original heading is
retained in `records.csv`. A personal name alone does not assign this category.

The two packs retain the existing `0.8-draft` weighted rules. They share category
IDs, but their terms and weights were curated independently. They are not the
English-reference concept packs of the related
[analysis repository](https://github.com/Kostiantyn-code/archival-description-analysis).
**This comparison tool does not yet classify English, Croatian or Italian titles.**
An English report does not imply an English or cross-language semantic classifier.

## Counts and denominators

An analyzable record has status `case` and a nonempty title. Dataset category
shares divide category counts by all analyzable titles in that workbook, including
unclassified titles. Inventory tables provide both inventory and dataset shares;
stacked category-chart segments use the workbook denominator. Multiple categories
can apply to one title, so their shares can sum to more than 100%.

Classification coverage has three exclusive states: subject classified;
context only (no subject category); unclassified (neither subject nor context).
Their counts sum to the number of analyzable titles, and shares sum to 100% apart
from rounding. Empty coverage tables use zero shares; document-form and network
rates instead use blank values when their denominator is absent.

With `chronology.expansion_mode: span`, a dated case contributes to every year of
its inclusive range, up to `maximum_span_years`; otherwise only its start year is
used. A range exceeding that maximum also contributes only to its start year.
The aggregate chronology shows counts. Each overlaid inventory series divides
annual counts by all analyzable titles of that inventory; annual shares can sum
to more than 100% because a case can span several years.

## Vocabulary and title similarity

Vocabulary counts document frequency: a stem counts at most once per title.
Both language packs' stopword stems are excluded, as are numbers and very short
words. Frequent terms are ranked by their title share. Distinctive terms use
`share × (log((1 + number_of_datasets)/(1 + datasets_containing_term)) + 1)`;
"containing" uses the configured minimum document frequency. Displayed terms are
the most frequent observed surface form of a stem. Shared vocabulary requires
the minimum frequency in every dataset. These are lexical comparisons, not
translation or aligned multilingual concept frequencies.

Similarity uses character-within-word TF-IDF n-grams (default 3–5), `min_df=2`,
sublinear term frequency and cosine nearest neighbours. For each dataset pair,
titles from the smaller dataset query the larger one. Up to `matches_per_case`
neighbours above `minimum_score` are retained for each query; the final table is
limited to 2,000 pairs. A corpus with no shared retained n-grams yields an empty
table. Similarity is not symmetric exhaustive pair retrieval and does not establish
identity, historical equivalence or translation between languages.

## Document forms and theme co-classification

Fifteen document-form regexes are maintained alongside comparative profiles in
`src/comparative_profiles.py`. Patterns originate from `archival-description-analysis`
commit `c4a772c`, rules version `0.1-draft`; consolidation changes their labels to
English, not their patterns. Each form counts once per title, independently of
subject categories and worksheet language. Generic “case concerning” headings do
not create a type. `document_type_mentions` retains matched wording and source
coordinates. A mention does not verify which documents are inside the case.

For each dataset and inventory, type rates use all analyzable titles:
`percent = mentions/N × 100`; `per_1000 = mentions/N × 1000`. Zero indicates that
there are titles but no match. With no titles, counts are zero and rates blank.

Only assigned subject categories form network pairs; context-only categories
are excluded. Categories are deduplicated, then mapped to distinct macroblocks.
Three categories create three unordered pairs; two categories of the same block
do not create a block self-loop. For N titles, marginals A and B, and shared S:

- shared frequency per 1,000 = `S/N × 1000`;
- lift = `S × N/(A × B)`.

For N=100, A=20, B=25 and S=10, the rate is 100 per 1,000 and lift is 2.
If A and B are nonzero but S=0, lift is zero. If A or B is zero, lift is undefined
(blank). All pairs, including zero-count pairs, are exported. Inventory counts
reconcile with dataset counts; lift is recomputed at each level, not averaged.
`theme_link_cases` identifies the contributing physical rows.

The offline network uses fixed node positions and shared metric scales. It shows
up to 24 category pairs or six block pairs. Lift view additionally requires
`max(3, ceil(N × 0.0015))` shared titles. These filters affect display only.
Examples show the first three contributing cases; complete evidence is in tables.
Lift is descriptive, not a significance test or evidence of causality; rare pairs
can have unstable high lift. Normalizing rates does not remove differences in
description detail, chronology, creator functions or dictionary coverage.

## Presentation, provenance and validation

CLI messages are English. Reports default to English; Ukrainian is selectable through
`visualization.language`. CSV field names and generated labels remain English
regardless of report language. XLSX/HTML captions and figures follow the setting.
Dataset/archive names are source values unless an explicit `label_overrides`
translation is supplied. UI templates contain explicit caption markers, rendered
before source values are embedded; titles are never passed through a translation
replacement. The network escapes JSON `<` characters and HTML table values are
escaped. Text cells in exported XLSX are stored as text, including titles starting
with `=`. Installation downloads dependencies; analysis and generated reports
require no remote services.

`run_manifest.json` records input SHA-256 hashes, analytical and presentation
source hashes, dictionary/rule versions, Python/dependency versions, report-language
settings and output paths. Keep the exact input workbooks, source checkout and
configuration together for reproduction. Requirements allow compatible version
ranges; the manifest records the actual installed versions.

The cleanup baseline is main commit `901f451`. Regression checks cover unchanged
rules and 3,900 deterministic UK/RU decisions derived from dictionary terms,
phrases, co-occurrence rules and ambiguity contexts, both alone and with a
correspondence prefix. A separate three-workbook/four-inventory trial includes
empty data, shifted rows, duplicate numbers, mixed archives and unsafe HTML text:
17 records, eight similar pairs and eight validation messages. Counts, category
sets/scores, evidence terms and analytical metrics are compared independently of
translated labels. Tests also run complete English/Ukrainian/bilingual reports,
check CSV invariance and parse generated network JavaScript. These checks establish
regression parity, **not independently measured linguistic accuracy**.
