# 90-second portfolio walkthrough — invented data only

**Purpose:** Show the thought process behind the data-quality work, not claim to compete with Partly's OEM/AI infrastructure.

## What to show

| Approx. time | Screen or proof | What to say |
| --- | --- | --- |
| 0–15 s | A line of invented supplier shorthand | “In parts retail, one line can mix model, year and multiple applications. I wanted to make it easier to inspect accurately.” |
| 15–35 s | Local Streamlit app → download and upload the built-in synthetic PDF | “This version extracts a text PDF and proposes normalized records, preserving the original line.” |
| 35–55 s | Application rows, review queue, and evidence/conflict sheets | “Uncertain information stays reviewable; it is not a verified part-to-vehicle fitment.” |
| 55–75 s | [GitHub Actions result](https://github.com/errelmaxino/pricelist-converter-app/actions/runs/43189215988) and quality scorecard | “We found real errors through synthetic tests: shorthand years and record-to-evidence links. We fixed those and kept regression checks.” |
| 75–90 s | Optional AI tab, source quote and explicit approval UI | “AI can suggest a correction, but it cannot replace a trusted source or human verification. The app never silently applies a change.” |

## Safe demo sequence

1. Run `streamlit run streamlit_app.py` locally using the README instructions.
2. Click **Download synthetic sample PDF**. Re-upload that invented file to the PDF input; **do not upload confidential supplier records for a public recording**.
3. Show normalized applications, one unresolved record, the evidence row key, and the workbook tabs.
4. If demonstrating the AI tab, use an entirely invented row and a manually prepared JSON suggestion. Disclose that it is a **manual proposal exchange**, not an autonomous AI API call.
5. State the limitation plainly: **synthetic record agreement is not real-world fitment accuracy**.

This walkthrough is a script; **no recorded video or deployed demonstration has been verified**.

## One-line résumé talking point (only after privacy/owner clearance)

“Directed an AI-assisted automotive catalog data-quality project with traceable PDF normalization, human review, and synthetic regression tests; corrected parsing and evidence-linking defects found during quality evaluation.”

## Useful follow-up questions

- How would you detect a missing parts application, not just a wrong one?
- Which fields need a manufacturer-authoritative reference before accepting a fitment?
- How would you test generalization across an unfamiliar supplier catalog?
- How would you review the mistakes a language model makes without giving it the final say?
