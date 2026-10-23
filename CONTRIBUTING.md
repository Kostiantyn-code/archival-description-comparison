# Contributing

Open an Issue or Pull Request with the problem, expected behaviour and a minimal
example. For input problems, include the Python version, operating system,
command and a small anonymized workbook layout. Do not publish restricted records
or personal data in examples.

Run before submitting:

```bash
python -m compileall -q main.py src tests
python -m unittest discover -s tests -v
```

For dictionary changes, give the language, category ID, example title and expected
result. Preserve weighted aliases unless their removal has been evaluated: two
similar forms can contribute separate scores. Update `dictionary_version` in
`config/categories.yaml` when rules change. Describe format changes in
`CHANGELOG.md`; see [methodology](docs/methodology.md).
