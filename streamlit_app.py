from __future__ import annotations

import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from parser import (
    MAIN_COLUMNS,
    ParsedCatalog,
    applications_dataframe,
    build_workbook,
    parse_catalog_pdf,
    raw_rows_dataframe,
    workbook_to_bytes,
)
from storage import export_profiles_json, get_profile, import_profiles_json, list_profiles, save_profile

st.set_page_config(page_title="Smart Automotive Pricelist PDF to Excel", page_icon="🧩", layout="wide")

st.title("Smart Automotive Pricelist PDF to Excel")
st.caption(
    "Split one catalog line into multiple compatible application rows, spot obvious patterns, and review only the rows that need attention."
)

with st.sidebar:
    st.subheader("Supplier profiles")
    profiles = list_profiles()
    st.caption(f"Saved profiles: {len(profiles)}")
    for profile in profiles[:8]:
        st.write(f"- **{profile['supplier_name']}** · updated {profile['updated_at']}")

    st.divider()
    st.subheader("Built-in pattern spotting")
    st.markdown(
        "- `TOY` → **TOYOTA**  \n"
        "- `MIT` / `MITS` → **MITSUBISHI**  \n"
        "- `NIS` → **NISSAN**  \n"
        "- `07 / 07' / '07 / \"07` → **2007** when clearly used as a year  \n"
        "- carries the brand forward across split models like `LANCER 07'` after `MIT. MIRAGE G4 14'`"
    )

    st.divider()
    st.subheader("Profile backup")
    st.download_button(
        "Download supplier profile backup",
        data=export_profiles_json().encode("utf-8"),
        file_name="supplier_profiles_backup.json",
        mime="application/json",
    )
    import_file = st.file_uploader("Import supplier profile backup", type=["json"], key="profile_import")
    if import_file is not None:
        try:
            imported = import_profiles_json(import_file.getvalue().decode("utf-8"))
            st.success(f"Imported {imported} supplier profile(s).")
        except Exception as exc:  # noqa: BLE001
            st.error(f"Could not import backup: {exc}")

supplier_name = st.text_input("Supplier name", value="NUVO / Nuvo-Pro")
profile_notes = st.text_area(
    "Supplier notes (optional)",
    value=(get_profile(supplier_name) or {}).get("notes", "") if supplier_name.strip() else "",
    height=90,
    placeholder="Example: Nuvo catalogs usually use TOY / MIT / NIS abbreviations and split many applications with '/'.",
)

uploaded = st.file_uploader("Upload automotive pricelist PDF", type=["pdf"])

if uploaded is not None:
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = Path(tmpdir) / uploaded.name
        pdf_path.write_bytes(uploaded.getvalue())
        with st.spinner("Reading catalog and splitting applications..."):
            parsed: ParsedCatalog = parse_catalog_pdf(pdf_path)

    applications_df = applications_dataframe(parsed.application_rows)
    raw_df = raw_rows_dataframe(parsed.raw_rows)
    review_df = pd.DataFrame(parsed.review_rows, columns=MAIN_COLUMNS)

    st.subheader("Conversion summary")
    metric_cols = st.columns(5)
    metric_cols[0].metric("Catalog rows", len(raw_df))
    metric_cols[1].metric("Application rows", len(applications_df))
    metric_cols[2].metric("Auto-accepted", int((applications_df["Review Status"] == "Auto-Accepted").sum()))
    metric_cols[3].metric("Review suggested", int((applications_df["Review Status"] == "Review Suggested").sum()))
    metric_cols[4].metric("Needs review", int((applications_df["Review Status"] == "Needs Review").sum()))

    if parsed.warnings:
        for warning in parsed.warnings:
            st.warning(warning)

    st.info(
        "The app now expands one catalog line into multiple application rows. For example, one line like `MIT. MIRAGE G4 14'/LANCER 07'/NISSAN CUBE 1.5` becomes separate rows for Mirage G4 2014, Lancer 2007, and Nissan Cube 1.5."
    )

    tabs = st.tabs(["1) Split preview", "2) Review dashboard", "3) Row detail", "4) Raw catalog rows", "5) Export"])

    with tabs[0]:
        st.subheader("Split application preview")
        filter_cols = st.columns([1, 1, 1, 1])
        status_filter = filter_cols[0].selectbox(
            "Show rows",
            options=["All", "Auto-Accepted", "Review Suggested", "Needs Review"],
            index=0,
        )
        code_filter = filter_cols[1].text_input("Filter by code", placeholder="Example: VKX-1267")
        brand_filter = filter_cols[2].text_input("Filter by vehicle brand", placeholder="Example: TOYOTA")
        page_filter = filter_cols[3].text_input("Filter by page", placeholder="Example: 13")

        preview_df = applications_df.copy()
        if status_filter != "All":
            preview_df = preview_df[preview_df["Review Status"] == status_filter]
        if code_filter.strip():
            preview_df = preview_df[preview_df["Code"].str.contains(code_filter.strip(), case=False, na=False)]
        if brand_filter.strip():
            preview_df = preview_df[preview_df["Brand"].str.contains(brand_filter.strip(), case=False, na=False)]
        if page_filter.strip().isdigit():
            preview_df = preview_df[preview_df["Page"].astype(str) == page_filter.strip()]

        st.caption("You can directly edit the final export values here before downloading the workbook.")
        edited_df = st.data_editor(
            preview_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "Confidence": st.column_config.ProgressColumn("Confidence", min_value=0.0, max_value=1.0, format="%.0f%%"),
                "Original Price (PHP)": st.column_config.NumberColumn("Original Price (PHP)", format="₱ %.2f"),
                "Your Price (PHP)": st.column_config.NumberColumn("Your Price (PHP)", format="₱ %.2f"),
                "Use Price (PHP)": st.column_config.NumberColumn("Use Price (PHP)", format="₱ %.2f"),
            },
            disabled=["Source Line", "Supplier Brand", "Category", "Code", "Original Price (PHP)", "Page", "Confidence", "Review Status", "Pattern Notes"],
            key="applications_editor",
        )
        st.session_state["edited_applications_df"] = edited_df

    with tabs[1]:
        st.subheader("Review dashboard")
        st.caption("Start here after the split preview. This page helps you jump straight to the rows that still need human review.")

        dash_cols = st.columns(4)
        dash_cols[0].metric("Needs review", int((applications_df["Review Status"] == "Needs Review").sum()))
        dash_cols[1].metric("Review suggested", int((applications_df["Review Status"] == "Review Suggested").sum()))
        dash_cols[2].metric("Missing brand", int((applications_df["Brand"].fillna("") == "").sum()))
        dash_cols[3].metric("Missing year", int((applications_df["Year"].fillna("") == "").sum()))

        strict_view = st.selectbox(
            "Queue filter",
            options=["All review rows", "Needs Review only", "Review Suggested only", "Missing brand or model", "Missing year", "Missing engine"],
            index=0,
        )

        queue_df = applications_df[applications_df["Review Status"] != "Auto-Accepted"].copy()
        if strict_view == "Needs Review only":
            queue_df = queue_df[queue_df["Review Status"] == "Needs Review"]
        elif strict_view == "Review Suggested only":
            queue_df = queue_df[queue_df["Review Status"] == "Review Suggested"]
        elif strict_view == "Missing brand or model":
            queue_df = queue_df[(queue_df["Brand"].fillna("") == "") | (queue_df["Model"].fillna("") == "")]
        elif strict_view == "Missing year":
            queue_df = queue_df[queue_df["Year"].fillna("") == ""]
        elif strict_view == "Missing engine":
            queue_df = queue_df[queue_df["Engine"].fillna("") == ""]

        st.dataframe(queue_df, use_container_width=True, hide_index=True)

    with tabs[2]:
        st.subheader("Row detail")
        st.caption("Inspect one split application row with the raw source line, confidence, and pattern notes.")
        detail_source_df = applications_df[applications_df["Review Status"] != "Auto-Accepted"].copy()
        if detail_source_df.empty:
            detail_source_df = applications_df.copy()

        detail_label = detail_source_df.apply(
            lambda r: f"{r['Code']} · {r['Brand'] or '(blank brand)'} · {r['Model'] or '(blank model)'} · p{r['Page']}",
            axis=1,
        ).tolist()
        detail_map = dict(zip(detail_label, detail_source_df.index.tolist()))
        selected_label = st.selectbox("Inspect row", options=detail_label)
        selected_row = detail_source_df.loc[detail_map[selected_label]]

        left, center, right = st.columns([1.1, 1.3, 1.0])
        with left:
            st.markdown("**Raw source line**")
            st.code(str(selected_row["Source Line"]))
            st.markdown("**Key parsed fields**")
            st.write(selected_row[["Code", "Axle", "Brand", "Model", "Engine", "Year"]])

        with center:
            st.markdown("**Current export values**")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Field": ["Supplier Brand", "Category", "Code", "Axle", "Brand", "Model", "Engine", "Year", "Original Price (PHP)"],
                        "Value": [
                            selected_row["Supplier Brand"],
                            selected_row["Category"],
                            selected_row["Code"],
                            selected_row["Axle"],
                            selected_row["Brand"],
                            selected_row["Model"],
                            selected_row["Engine"],
                            selected_row["Year"],
                            selected_row["Original Price (PHP)"],
                        ],
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

        with right:
            st.markdown("**Why this row was scored this way**")
            st.write(f"Status: **{selected_row['Review Status']}**")
            st.write(f"Confidence: **{int(float(selected_row['Confidence']) * 100)}%**")
            st.write(f"Pattern notes: {selected_row['Pattern Notes'] or '—'}")
            st.write(
                "High-confidence rows are auto-accepted. Medium-confidence rows go into Review Suggested. Missing key values or weak matches go into Needs Review."
            )

    with tabs[3]:
        st.subheader("Raw catalog rows before splitting")
        st.caption("This shows the original line-level rows detected from the PDF before the app expands them into compatible vehicle rows.")
        st.dataframe(raw_df, use_container_width=True, hide_index=True)

    with tabs[4]:
        st.subheader("Export")
        final_applications_df = st.session_state.get("edited_applications_df", applications_df)
        review_export_df = final_applications_df[final_applications_df["Review Status"] != "Auto-Accepted"].copy()

        export_cols = st.columns([1, 1])
        with export_cols[0]:
            if st.button("Save / update supplier profile"):
                try:
                    save_profile(supplier_name=supplier_name, notes=profile_notes)
                    st.success(f"Saved supplier profile for {supplier_name.strip()}.")
                except Exception as exc:  # noqa: BLE001
                    st.error(str(exc))
        with export_cols[1]:
            st.write("")

        workbook = build_workbook(final_applications_df, raw_df, review_export_df)
        excel_bytes = workbook_to_bytes(workbook)
        output_name = f"{Path(uploaded.name).stem}_smart_split.xlsx"
        st.download_button(
            "Download Excel workbook",
            data=excel_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        st.markdown("**What is inside the workbook**")
        st.markdown(
            "- **Applications**: final split rows ready for pricing and cleanup  \n"
            "- **Catalog Rows**: original detected line-level rows  \n"
            "- **Review Queue**: only the rows that were not auto-accepted"
        )

        st.markdown("**Recommended flow**")
        st.markdown(
            "1. Upload PDF  \n"
            "2. Check the **Split preview**  \n"
            "3. Review only the **Review queue** rows  \n"
            "4. Edit anything needed in the preview table  \n"
            "5. Download the Excel workbook"
        )
else:
    st.subheader("Recommended app flow")
    flow_cols = st.columns(4)
    flow_cols[0].markdown("**1. Upload**  \nUpload the catalog PDF and choose the supplier.")
    flow_cols[1].markdown("**2. Auto-split**  \nThe parser turns one line into multiple brand/model application rows.")
    flow_cols[2].markdown("**3. Review**  \nOnly medium-confidence rows go into the review queue.")
    flow_cols[3].markdown("**4. Export**  \nDownload a workbook with Applications, Catalog Rows, and Review Queue sheets.")

    st.markdown("### Example of the split behavior")
    st.code("VKX-1267 REAR MIT. MIRAGE G4 14'/LANCER 07'/NISSAN CUBE 1.5")
    st.markdown(
        "Becomes three rows:  \n"
        "- Mitsubishi · Mirage G4 · 2014  \n"
        "- Mitsubishi · Lancer · 2007  \n"
        "- Nissan · Cube 1.5"
    )
