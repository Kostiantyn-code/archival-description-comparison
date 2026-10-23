# Changelog

## Unreleased

- Consolidate 26 dictionary fragments and two stopword lists into `uk.yaml` and
  `ru.yaml`, preserving weighted rules and category IDs.
- Default to English, including CLI messages, workbook/HTML captions and generated
  analytical labels; retain Ukrainian reports and bilingual figure labels.
- Use explicit UI markers in HTML/JS templates, separate charts from table export,
  and keep document-form rules with comparative profiles.
- Recognize English and Russian worksheet headers/metadata; return empty similarity
  tables when no shared n-grams survive; preserve formula-like XLSX strings as text.
- Record analytical source hashes and installed dependency versions in manifests.
- Merge the two READMEs into a shorter English-first/Ukrainian guide and consolidate
  dictionary/audit documentation in the English methodology document.
- Add regression checks against the pre-cleanup rules and 3,900 UK/RU decisions.
- Include 15 document-form indicators and category/macroblock co-classification
  with counts, rates per 1,000, lift and source evidence. The offline network uses
  shared scales, dataset/inventory selection and cross-dataset comparison.
- Use physical Excel rows in record IDs, retaining repeated case numbers separately.
- Support personnel-file section inheritance and explicit classification evidence.
- Add per-inventory chronology, coverage and category tables. Category panels share
  a scale, with inventory segments and one grouped legend entry per workbook.
- Export all six static figures as PNG and editable SVG.
- Group records before summary counts; share YAML/header/classification helpers.
- Reuse Windows Python environments between ZIP extractions; both launchers update
  dependencies only after requirements change.

## 0.2.0 — 2026-09-22

- Prepare the first public standalone repository with bilingual documentation,
  contribution guidance, citation metadata, an MIT license and issue templates.
- Add Linux/macOS and improved Windows launchers.
- Add GitHub Actions for Python 3.10, 3.12 and 3.13.
- Replace dataset configuration with a neutral public example.
- Allow input workbooks outside the project directory in run manifests.
- Keep the software version in one module.

## 0.1.1

- Make each dataset's analyzable titles the explicit denominator for category shares.
- Add classification coverage and separate subject-unclassified/context-only lists.
- Test percentage formulas and mutually exclusive classification states.

## 0.1.0

- Initial local comparison of two or more multi-sheet XLSX workbooks.
- Read archive/fond/inventory metadata per worksheet and recover supported shifted rows.
- Classify titles with version 0.8 dictionaries; compare chronology, themes,
  vocabulary and cross-dataset title similarity.
- Export XLSX, CSV, HTML, PNG and validation logs.
