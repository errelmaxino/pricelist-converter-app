# Automotive catalog project — quality and employer-portfolio audit

**Date:** 2026-11-12 (Philippine time)  
**Source of truth:** GitHub `main` `0199c2e9` and draft [PR #2](https://github.com/errelmaxino/pricelist-converter-app/pull/2).  
**Goal:** Show traceable catalog normalization and practical automotive data-quality thinking. This is **not** a certified fitment lookup.

## Scope and method

Inspected `parser.py`, `streamlit_app.py`, `storage.py`, `requirements.txt`, developer-container settings, existing security tests, PR diff, GitHub Actions checks/logs, and employer-facing README/case study. Added synthetic tests and exercised them via GitHub CI on Python 3.11 and 3.12. The public binary `demo_output.xlsx` could not be inspected through the available GitHub file tools; its provenance and redistribution permission remain **UNKNOWN**. Did not run the Streamlit app in an authenticated browser or test any production deployment.

## Findings ranked by effect on trust

| Severity | Finding | Evidence | Status |
| --- | --- | --- | --- |
| High | Review UI generated a different evidence/conflict key from the parser, so the supporting records could be linked to the wrong application row | `_row_key_for_idx` vs `expand_catalog_rows` | **Fixed on draft branch**; source row identity is now exported as `Row Key` and joined to evidence |
| High | Repeated part codes on one PDF page generated duplicate source keys | Former `P{page}|{code}|{segment}` | **Fixed on draft branch** using page, raw-row ordinal, code, segment; synthetic duplicate-code tests |
| High | Same-name, same-length but different PDF files shared the editor working-state signature | `_signature` previously included byte count but not file contents | **Fixed on draft branch** with SHA-256 of source bytes; regression tests |
| High | Existing public `demo_output.xlsx` could contain supplier information or unlicensed content | Committed binary already on `main`; not readable via this audit surface | **UNKNOWN / portfolio-sharing hold**; owner/source-rights review required, and Git history remains public |
| Medium | A quoted short year left stray punctuation in the vehicle model, such as `COROLLA '` | Newly added synthetic test initially failed | **Fixed**, regression tested on two Python versions |
| Medium | `Confidence` and `Auto-Accepted` are rule-derived states, not calibrated rates or manufacturer fitment confirmations | `parser.py`, Streamlit bulk controls | **Not solved**: UI and documentation now warn explicitly; do not market confidence as measured truth |
| Medium | Originally no end-to-end synthetic PDF golden-set check | Previously security-focused tests only | **Improved**: 8 invented PDF lines, 9 expected applications, exact record matching and evidence lineage checks |
| Medium | Developer-container preview turned off CORS and XSRF protections | `.devcontainer/devcontainer.json` | **Removed unsafe overrides on draft branch**; developer preview still needs practical acceptance |
| Medium | Hard-coded layout and supported part number prefixes constrain real-supplier coverage; scanned PDFs unsupported | `extract_catalog_rows`, `CODE_REGEX`, PDF extraction | **Known limitation**; test with permission-cleared supplier samples before claiming wider accuracy |
| Low | AI assistance was in the development process, not used at runtime | Previously rule-based app | **Improved**: optional provider-neutral copy/paste AI proposal workflow; exact-source and human approval gates; no automatic model API |
| Medium | Production visual, deployed runtime, browser accessibility, and real supplier accuracy are untested | No verified deployment/session evidence in this audit | **Pending**; don't claim production readiness |

## What the synthetic benchmark actually proves

`python scripts/quality_scorecard.py` generates an **invented text PDF** using only repository text and standard Python. The expected answers are written independently in `tests/synthetic_catalog_fixture.py`.

The scorecard reports extracted catalog lines, generated application records, exact record matches on `(part code, brand, model, year)`, precision/recall on **synthetic records only**, review cases, and uniqueness of source row IDs.

This does **not** prove OEM fitment, price truth, real-world generalization, catalog completeness, or 100% employer-data accuracy. A nontrivial real benchmark needs an authorized, independently labeled set of supplier pages and a way to count omissions as well as wrong records.

## AI strategy and guardrails

**Implemented MVP:** `AI review (opt-in)` tab. One selected row becomes a provider-neutral prompt. The user decides if/where to send it. Pasted JSON must match the row key and exact source line, include an exact evidence excerpt for proposed changes, and use a strict field allowlist. No API call or charge occurs in the app. No suggestion is silently applied; a human must explicitly inspect and approve it; the row remains `Needs Review`.

**Not implemented:** autonomous paid AI API, provider keys, external OEM lookups, validation against authoritative fitment databases, calibrated AI accuracy, or real-time inventory sync. Even valid JSON and a quoted excerpt do **not** guarantee the suggestion is semantically correct.

**Next reversible test:** use only a permission-cleared or synthetic labeled sample to compare rule-only output, AI suggestion output, and human-corrected gold answers. Add an API only if it reduces real errors or labor enough to justify spend, privacy, and maintenance.

## Public portfolio acceptance gate

Before an employer-facing link, require:
1. Clear `demo_output.xlsx` and any other existing public artifact for confidentiality, license, prices, and reuse rights; replacing on the default branch does **not** erase Git history.
2. Independently review PR claims, diff, security, and tests; fix material problems.
3. Verify the Streamlit UI with the safe sample PDF and the exported workbook at realistic screen size.
4. Obtain the owner's approval of project-attribution wording and application language.
5. Merge only after the public-data check and acceptance; don't imply the unmerged PR is a live product.

## Portfolio relevance

The hiring-value signal is **domain expertise + initiative + falsifiable quality work**. The project demonstrates recognition of messy catalog sources, safe normalization, source-traceable proposals, human exception handling, and corrections driven by tests. It does not imply professional software-engineering tenure or a full vehicle-fitment solution.
