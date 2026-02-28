from __future__ import annotations

import json
import tempfile
from pathlib import Path

import streamlit as st

from parser import (
    CANONICAL_COLUMNS,
    IGNORE_COLUMN,
    KEEP_AS_EXTRA,
    apply_mapping,
    build_suggested_mapping,
    build_workbook,
    dataframe_preview,
    extract_document,
    workbook_to_bytes,
)
from storage import export_templates_json, find_template, import_templates_json, list_templates, save_template

st.set_page_config(page_title="Multi-Format Pricelist PDF to Excel", page_icon="📄", layout="wide")

st.title("Multi-Format Pricelist PDF to Excel")
st.caption("Upload a PDF, auto-detect columns, map them once, and reuse the template for the same supplier later.")

with st.sidebar:
    st.subheader("Templates")
    templates = list_templates()
    st.caption(f"Saved templates: {len(templates)}")
    if templates:
        for tpl in templates[:8]:
            st.write(f"- **{tpl['supplier_name']}** · {tpl['mode']} · {tpl['header_fingerprint']}")
    exported = export_templates_json()
    st.download_button(
        "Download template backup",
        data=exported.encode("utf-8"),
        file_name="supplier_templates.json",
        mime="application/json",
    )
    import_file = st.file_uploader("Import template backup", type=["json"], key="import_json")
    if import_file is not None:
        try:
            count = import_templates_json(import_file.getvalue().decode("utf-8"))
            st.success(f"Imported {count} template(s).")
        except Exception as exc:
            st.error(f"Import failed: {exc}")

with st.expander("What this version can do", expanded=True):
    st.markdown(
        """
- **Preserves unknown columns** instead of dropping them
- **Auto-detects headers** from many text-based pricelist PDFs
- Lets you **map detected columns once** to your preferred output
- **Remembers supplier formats** using saved templates
- Supports two output styles:
  - **Normalized pricelist** - your standard columns + extra columns preserved
  - **Exact detected columns** - close to the uploaded file's own columns
        """
    )
    st.info(
        "Saved templates use a local database inside the app package. On some free cloud hosts, that storage can reset after a rebuild."
        " Use the template backup download if you want a durable copy."
    )

supplier_name = st.text_input("Supplier name", placeholder="Example: FIC / Supplier A")
output_mode = st.radio(
    "Output style",
    options=["normalized", "exact"],
    format_func=lambda x: "Normalized pricelist" if x == "normalized" else "Exact detected columns",
    horizontal=True,
)
preserve_unknown = st.checkbox("Preserve unmapped columns as extra Excel columns", value=True)
uploaded = st.file_uploader("Choose a PDF pricelist", type=["pdf"])

if uploaded is not None:
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = Path(tmpdir) / uploaded.name
        pdf_path.write_bytes(uploaded.getvalue())
        with st.spinner("Reading PDF and detecting tables..."):
            doc = extract_document(pdf_path)

    st.write(f"**Header fingerprint:** `{doc.header_fingerprint}`")
    if doc.warnings:
        for warning in doc.warnings:
            st.warning(warning)

    matched_template = find_template(supplier_name=supplier_name, header_fingerprint=doc.header_fingerprint)
    suggested_mapping = build_suggested_mapping(doc.detected_headers)
    state_key = f"mapping::{doc.header_fingerprint}::{supplier_name or 'unknown'}"

    if state_key not in st.session_state:
        if matched_template:
            st.session_state[state_key] = matched_template["mapping"]
            st.session_state[f"mode::{state_key}"] = matched_template.get("mode", output_mode)
            st.session_state[f"preserve::{state_key}"] = matched_template.get("preserve_unknown", preserve_unknown)
        else:
            st.session_state[state_key] = suggested_mapping
            st.session_state[f"mode::{state_key}"] = output_mode
            st.session_state[f"preserve::{state_key}"] = preserve_unknown

    if matched_template:
        st.success(
            f"Applied saved template from supplier **{matched_template['supplier_name']}**"
            f" (updated {matched_template['updated_at']})."
        )

    output_mode = st.radio(
        "Current file output style",
        options=["normalized", "exact"],
        key=f"mode::{state_key}",
        format_func=lambda x: "Normalized pricelist" if x == "normalized" else "Exact detected columns",
        horizontal=True,
    )
    preserve_unknown = st.checkbox(
        "Preserve unmapped columns as extra Excel columns",
        key=f"preserve::{state_key}",
        value=st.session_state[f"preserve::{state_key}"] if f"preserve::{state_key}" in st.session_state else True,
    )

    st.subheader("Detected columns")
    st.caption("Adjust mapping if needed. Anything set to 'Keep as extra column' will still appear in the Excel file.")
    options = CANONICAL_COLUMNS[:-2] + [KEEP_AS_EXTRA, IGNORE_COLUMN]
    labels = {KEEP_AS_EXTRA: "Keep as extra column", IGNORE_COLUMN: "Ignore"}

    cols = st.columns(2)
    current_mapping = st.session_state[state_key]
    for idx, header in enumerate(doc.detected_headers):
        column = cols[idx % 2]
        default_target = current_mapping.get(header, suggested_mapping.get(header, KEEP_AS_EXTRA))
        option_index = options.index(default_target) if default_target in options else options.index(KEEP_AS_EXTRA)
        selection = column.selectbox(
            header,
            options=options,
            index=option_index,
            format_func=lambda x: labels.get(x, x),
            key=f"map::{state_key}::{header}",
        )
        current_mapping[header] = selection

    final_headers, final_rows = apply_mapping(doc, current_mapping, mode=output_mode, preserve_unknown=preserve_unknown)
    preview_df = dataframe_preview(final_headers, final_rows, limit=25)

    st.subheader("Preview")
    st.dataframe(preview_df, use_container_width=True, hide_index=True)

    wb = build_workbook(final_headers, final_rows)
    excel_bytes = workbook_to_bytes(wb)
    out_name = f"{Path(uploaded.name).stem}_{output_mode}.xlsx"
    st.download_button(
        "Download Excel file",
        data=excel_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    save_cols = st.columns([2, 1])
    with save_cols[0]:
        if st.button("Save / update supplier template", type="secondary"):
            if not supplier_name.strip():
                st.error("Enter a supplier name first so the template can be saved and reused.")
            else:
                payload = {
                    "supplier_name": supplier_name.strip(),
                    "header_fingerprint": doc.header_fingerprint,
                    "mode": output_mode,
                    "preserve_unknown": preserve_unknown,
                    "mapping": current_mapping,
                }
                save_template(payload)
                st.success(f"Saved template for {supplier_name.strip()}.")
    with save_cols[1]:
        current_template_json = json.dumps(
            {
                "supplier_name": supplier_name.strip() or "sample-supplier",
                "header_fingerprint": doc.header_fingerprint,
                "mode": output_mode,
                "preserve_unknown": preserve_unknown,
                "mapping": current_mapping,
            },
            ensure_ascii=False,
            indent=2,
        )
        st.download_button(
            "Download current template",
            data=current_template_json.encode("utf-8"),
            file_name=f"template_{doc.header_fingerprint}.json",
            mime="application/json",
        )
