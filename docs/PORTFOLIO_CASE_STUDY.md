# Portfolio case study: Making messy automotive parts data reviewable

**Project:** [Automotive Parts Catalog Quality / PDF to Excel](../README.md)  
**Stage:** Working codebase with basic security regression tests; full production/deployed acceptance not proven  
**Focus:** Automotive parts catalog structure, normalization, discrepancy detection and human review

## 1. Why this problem matters

Parts counter staff and marketplace teams depend on a supplier's descriptions, stock records and catalog shorthand. An incorrect vehicle-year interpretation or missing model can waste time and create wrong-part inquiries. The aim here was not to invent a universal parts database, but to make messy source information easier to inspect and correct before re-use.

## 2. The specific challenge

A supplier entry may contain an abbreviation (`TOY`), a quoted short year (`'07`), several vehicle applications separated by slashes, or a model name that looks like a year (`BONGO 2000`). Flattening these into one untraceable row loses information; blindly guessing creates false confidence.

The application therefore keeps **two distinct levels**: the original extracted catalog row and proposed structured application rows.

## 3. What is implemented in the repository

- Page-aware PDF text extraction and source lines.
- Brand aliases, model phrases, year handling, position/engine parsing and application splitting.
- Field validation, per-field evidence with source/reason/confidence, conflict tracking and review queues.
- Manual correction and Excel output containing five separately named worksheets.
- Local supplier-specific parsing preferences and import/export of profiles.
- Five security-boundary tests for path traversal and spreadsheet formula injection, merged in PR #1.
- Four synthetic data-quality tests added in draft PR #2. These initially exposed a quoted-year defect: `'07` produced model `COROLLA '` instead of `COROLLA`. A narrowly scoped parser change now consumes the full year token without leaving stray punctuation.

See `parser.py`, `streamlit_app.py`, `storage.py`, and `tests/` for code and checks. These are feature descriptions observed in the code, not measured field-accuracy or customer outcomes.

## 4. Walkthrough with invented data

| Invented input | Why it's tricky | Intended handling |
| --- | --- | --- |
| `TOY COROLLA '07 / MIT LANCER '08` | Two brands and shorthand years in one description | Split candidates and normalize supported aliases/years, retaining source linkage |
| `MAZ BONGO 2000` | Number is part of the model name | Protect `BONGO 2000`; do not manufacture a 2000 model-year claim |
| `COROLLA 2007` | Brand isn't explicit in an isolated row | Leave it unresolved/reviewable rather than assert a specific brand |

The [synthetic regression checks](../tests/test_catalog_quality.py) and [generated PDF end-to-end tests](../tests/test_end_to_end_synthetic.py) pin down these decisions. The generated PDF has eight invented catalog lines and nine expected structured records; it is explicitly **not** a real supplier catalog or verified fitment dataset.

## 5. Decisions and tradeoffs

- **Trust the source over a pretty spreadsheet.** Preserve raw rows and page references.
- **Don't equate parsing with compatibility.** A plausible model/year candidate does not prove a part fits the vehicle.
- **Keep a human in control.** Low-confidence and conflicting information should remain reviewable.
- **Prefer explicit rules and visible evidence over an unverified AI claim.** The current parser is rule-based, even when its interface describes a proposal.
- **Protect exported data.** Keep untrusted supplier strings as literal text rather than executable spreadsheet formulas.
- **Make provenance inspectable.** Export stable row identifiers that match the evidence and conflict log even when part codes repeat.
- **Treat AI as a reviewer, not an authority.** The optional AI tab generates a one-row prompt, accepts source-matching JSON proposals and requires manual approval; no data is automatically sent to a model and nothing is silently applied.

## 6. Evidence and limits

**Verified check:** GitHub Actions [run 43185239223](https://github.com/errelmaxino/pricelist-converter-app/actions/runs/43185239223) completed successfully at branch commit `7deadb0`: **9/9 tests passed on Python 3.11 and 9/9 on Python 3.12**, with Python syntax checks. This includes five preexisting security tests and four new synthetic data-quality tests. The branch remains a draft PR, **not merged or deployed**. These tests do not establish real-world parts-fitment accuracy.

**New synthetic-only checkpoint:** The generated PDF's eight invented source lines yielded nine expected applications; all nine matched on (code, brand, model, year), one ambiguous brand remained flagged for review, and each application had a distinct traceable row key. A separate regression test also proved identical file names and byte counts do not cause edited state to carry between different PDFs. [GitHub CI run 43189215988](https://github.com/errelmaxino/pricelist-converter-app/actions/runs/43189215988) records the scorecard. Later checks continue on the draft PR.

**Not proven:** A measured reduction in errors or time, successful use across all real supplier catalogs, manufacturer/OEM fitment verification, public-hosted production performance, or direct synchronization with AMX.

The tracked `demo_output.xlsx` has not been cleared for disclosure; use only synthetic data for presentations and interviews until its rights and contents are reviewed.

## 7. What this says about the project owner

The core motivation came from real motorcycle/automotive retail operations: translating shorthand product lists, dealing with inventory record discrepancies, and helping customers identify compatible parts. The owner framed the workflow and directed **AI-assisted development**. This is a domain-led, learning-driven project—not a claim of personally authoring every line of code or having worked professionally as an OEM catalog data specialist.

## 8. Natural next step

The synthetic PDF, end-to-end tests and traceability checks are now implemented on draft PR #2. The [safe demo walkthrough](DEMO_WALKTHROUGH.md) describes an employer-friendly 90-second story. Next clear the already-public demo workbook's data/rights, independently review the changes, and manually verify Streamlit screens and exports before sharing externally. Only then consider real labeled catalogs or a paid AI endpoint.
