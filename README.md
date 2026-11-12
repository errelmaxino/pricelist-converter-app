# Automotive Parts Catalog Quality — PDF to Reviewable Data

**A practical, accuracy-first project for cleaning messy automotive supplier catalogs.**

[Read the case study](docs/PORTFOLIO_CASE_STUDY.md) · [See the synthetic data-quality checks](tests/test_catalog_quality.py)

## The problem

In parts retail, a single PDF catalog entry can refer to several vehicle models, contain abbreviated brand names, or mix a model name with a year and engine code. Copying it into a spreadsheet by hand is slow; treating every parsed string as a verified vehicle fitment is unsafe.

**Question this project explores:** How can we turn a supplier PDF into structured, inspectable records *while keeping uncertainty and source evidence visible*?

## What the existing application does

1. **Extract:** Read text-based automotive PDF catalogs and retain raw catalog lines and page numbers.
2. **Normalize:** Interpret supported brand aliases, shorthand years, positions, engine tokens, and slash-separated vehicle applications.
3. **Check:** Validate field formats, record field-level evidence and conflicts, and assign heuristic confidence scores.
4. **Review:** Show editable rows, missing important fields and low-confidence records for a human to inspect.
5. **Export:** Produce an Excel workbook with Applications, Catalog Rows, Review Queue, Conflict Log and Evidence Log sheets.

The code is Python + Streamlit, with pdfplumber for PDF extraction, pandas for tables, openpyxl for Excel export, and local SQLite supplier-specific parsing preferences.

## A small synthetic example

| Source description *(invented for illustration)* | Structured candidates |
| --- | --- |
| `TOY COROLLA '07 / MIT LANCER '08` | Toyota Corolla, 2007; Mitsubishi Lancer, 2008 |
| `MAZ BONGO 2000` | Mazda Bongo 2000 (model wording preserved; **not** assumed to mean model year 2000) |
| `COROLLA 2007` | Brand ambiguous; requires human review rather than inventing a fitment |

These are **synthetic rule examples**, not supplier records or proof of fitment. See the [automated data-quality tests](tests/test_catalog_quality.py). The parser's confidence measures rule extraction, **not** confirmation from an OEM/vehicle database.

## Try it

Python 3.11 is specified in `runtime.txt`.

```bash
python -m pip install -r requirements.txt
streamlit run streamlit_app.py
```

Upload a text-based PDF of the supported supplier layout, inspect the review and conflict tabs, and download the Excel output. There is no verified public hosted demo; this is a source repository and local-run workflow.

For repository checks:

```bash
python -m unittest discover -s tests -v
```

The synthetic tests exercise the existing normalization and review logic without proprietary supplier records. Existing security regression tests check upload-path confinement and literal Excel text serialization.

## What the project demonstrates

- Translating automotive parts catalog problems into explicit parsing and validation rules.
- Retaining original source references when transforming messy information.
- Separating **extracted**, **inferred**, **conflicted**, and **human-reviewed** information.
- Choosing a human review queue instead of claiming uncertain vehicle compatibility.
- Iterating safely: public security fixes were merged in [PR #1](https://github.com/errelmaxino/pricelist-converter-app/pull/1).
- Fixing a real catalog-cleaning defect: new synthetic tests exposed a stray apostrophe left in model names after quoted shorthand-year parsing. The parser was corrected, with [9/9 tests passing on Python 3.11 and 3.12](https://github.com/errelmaxino/pricelist-converter-app/actions/runs/43185239223) on the draft branch. No live fitment accuracy was claimed.

**Project approach:** The problem and product direction come from hands-on automotive parts retail and supplier-data workflows. Development is AI-assisted and owner-directed; this repository is not presented as proof of solo professional software engineering.

## Honest limitations

- **Not a complete Smart Parts Finder**: it does not verify whether a part fits a particular vehicle using OEM catalogs, VIN decoding, or authoritative cross-reference databases.
- Optimized for supported **text-based PDF** layouts; scanned/image-only PDFs require separate OCR that is **not implemented**.
- Confidence scores are heuristic extraction scores, **not fitment guarantees or independently validated accuracy rates**.
- Supplier-specific aliases and profiles are locally stored; there is no verified live data sync to AMX, retailer inventory, or external AI service.
- The existing `demo_output.xlsx` is a historical artifact. Its data provenance and permission for reuse in a public portfolio **have not been independently verified**; do not republish or treat it as an approved sample.
- Full deployed Streamlit acceptance, repeat-code row identity, and review/evidence join behavior remain unverified. The synthetic tests do not cover all supplier layouts.

## Next improvements (not yet implemented)

1. Verify and fix stable review-row/evidence identifiers when repeated supplier codes occur.
2. Add a clearly licensed synthetic PDF demo and a repeatable before/after quality check.
3. Require evidence-backed external cross-references before any **real fitment** claim.
4. Confirm the public workbook's redistribution rights and replace it with a synthetic fixture if needed.

See the [portfolio case study](docs/PORTFOLIO_CASE_STUDY.md) for the business context, design tradeoffs and proof boundaries.
