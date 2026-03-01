from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Iterable

import pandas as pd
import streamlit as st

from parser import (
    ParsedCatalog,
    applications_dataframe,
    build_workbook,
    conflict_dataframe,
    evidence_dataframe,
    parse_catalog_pdf,
    raw_rows_dataframe,
    workbook_to_bytes,
)
from storage import export_profiles_json, get_profile, import_profiles_json, list_profiles, save_profile

st.set_page_config(page_title="Smart Automotive Pricelist PDF to Excel", page_icon="🧩", layout="wide")


EDITABLE_FIELDS = [
    "Axle",
    "Side",
    "Vertical",
    "Mount",
    "Brand",
    "Model",
    "Engine",
    "Year",
    "Your Price (PHP)",
    "Use Price (PHP)",
]

REVIEW_REQUIRED_FIELDS = ["Brand", "Model", "Year", "Engine"]
LOW_CONFIDENCE_THRESHOLD = 0.80
AUTO_ACCEPT_THRESHOLD = 0.95
SAFE_PATTERN_TERMS = (
    "axle token",
    "side token",
    "vertical token",
    "mount token",
    "explicit brand",
    "short year range",
    "quoted short year",
    "single four-digit year",
    "trailing short year",
    "four-digit year range",
    "engine parsed",
)


def _parse_alias_text(text: str) -> dict[str, str]:
    aliases: dict[str, str] = {}
    for line in text.splitlines():
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()
        if key and value:
            aliases[key] = value
    return aliases


@st.cache_data(show_spinner=False)
def _cached_parse_catalog(
    file_name: str,
    file_bytes: bytes,
    custom_aliases: tuple[tuple[str, str], ...],
    protected_phrases: tuple[str, ...],
) -> ParsedCatalog:
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = Path(tmpdir) / file_name
        pdf_path.write_bytes(file_bytes)
        return parse_catalog_pdf(
            pdf_path,
            custom_aliases=dict(custom_aliases),
            protected_phrases=list(protected_phrases),
        )



def _signature(file_name: str, file_bytes: bytes, supplier_name: str, profile_notes: str, custom_aliases: dict[str, str], protected_phrases: list[str]) -> str:
    return "|".join(
        [
            file_name,
            str(len(file_bytes)),
            supplier_name.strip(),
            profile_notes.strip(),
            repr(sorted(custom_aliases.items())),
            repr(sorted(protected_phrases)),
        ]
    )



def _ensure_working_state(base_df: pd.DataFrame, signature: str) -> None:
    if st.session_state.get("working_signature") != signature:
        st.session_state["working_signature"] = signature
        st.session_state["working_df"] = base_df.copy()
        st.session_state["dashboard_queue_filter"] = "All review rows"
        st.session_state["dashboard_search_code"] = ""
        st.session_state["dashboard_search_brand"] = ""
        st.session_state["dashboard_selected_labels"] = []
        st.session_state["detail_pointer"] = 0



def _working_df() -> pd.DataFrame:
    return st.session_state["working_df"]



def _set_working_df(df: pd.DataFrame) -> None:
    st.session_state["working_df"] = df



def _safe_pattern_mask(df: pd.DataFrame) -> pd.Series:
    notes = df["Pattern Notes"].fillna("").astype(str).str.lower()
    mask = notes.apply(lambda x: any(term in x for term in SAFE_PATTERN_TERMS))
    mask &= ~notes.str.contains("brand carried forward", case=False, na=False)
    return mask



def _conflict_key_set(conflict_df: pd.DataFrame) -> set[str]:
    if conflict_df.empty:
        return set()
    return set(conflict_df["Row Key"].astype(str))



def _row_key_for_idx(df: pd.DataFrame, idx: int) -> str:
    row = df.loc[idx]
    return f"P{row['Page']}|{row['Code']}|{idx+1}"



def _important_blank_mask(df: pd.DataFrame) -> pd.Series:
    return (
        (df["Brand"].fillna("") == "")
        | (df["Model"].fillna("") == "")
        | (df["Year"].fillna("") == "")
        | (df["Engine"].fillna("") == "")
    )



def _queue_filter_options() -> list[str]:
    return [
        "All review rows",
        "Needs Review only",
        "Review Suggested only",
        "Conflicts only",
        "Blank important fields",
        f"Low confidence (<{int(LOW_CONFIDENCE_THRESHOLD * 100)}%)",
        "Missing brand or model",
        "Missing year",
        "Missing engine",
        "Carried brand rows",
        "Rows mentioned in conflict log",
    ]



def _apply_queue_filter(df: pd.DataFrame, filter_name: str, conflict_df: pd.DataFrame) -> pd.DataFrame:
    queue_df = df.copy()
    if filter_name == "All review rows":
        queue_df = queue_df[queue_df["Review Status"] != "Auto-Accepted"]
    elif filter_name == "Needs Review only":
        queue_df = queue_df[queue_df["Review Status"] == "Needs Review"]
    elif filter_name == "Review Suggested only":
        queue_df = queue_df[queue_df["Review Status"] == "Review Suggested"]
    elif filter_name == "Conflicts only":
        keys = _conflict_key_set(conflict_df)
        queue_df = queue_df[queue_df.index.map(lambda i: _row_key_for_idx(df, i) in keys)]
    elif filter_name == "Blank important fields":
        queue_df = queue_df[_important_blank_mask(queue_df)]
    elif filter_name.startswith("Low confidence"):
        queue_df = queue_df[queue_df["Confidence"].fillna(0.0).astype(float) < LOW_CONFIDENCE_THRESHOLD]
    elif filter_name == "Missing brand or model":
        queue_df = queue_df[(queue_df["Brand"].fillna("") == "") | (queue_df["Model"].fillna("") == "")]
    elif filter_name == "Missing year":
        queue_df = queue_df[queue_df["Year"].fillna("") == ""]
    elif filter_name == "Missing engine":
        queue_df = queue_df[queue_df["Engine"].fillna("") == ""]
    elif filter_name == "Carried brand rows":
        queue_df = queue_df[queue_df["Pattern Notes"].fillna("").astype(str).str.contains("brand carried forward", case=False, na=False)]
    elif filter_name == "Rows mentioned in conflict log":
        keys = _conflict_key_set(conflict_df)
        queue_df = queue_df[queue_df.index.map(lambda i: _row_key_for_idx(df, i) in keys)]
    return queue_df



def _apply_search_filters(df: pd.DataFrame, code_search: str, brand_search: str) -> pd.DataFrame:
    filtered = df.copy()
    if code_search.strip():
        filtered = filtered[filtered["Code"].fillna("").astype(str).str.contains(code_search.strip(), case=False, na=False)]
    if brand_search.strip():
        filtered = filtered[filtered["Brand"].fillna("").astype(str).str.contains(brand_search.strip(), case=False, na=False)]
    return filtered



def _merge_editor_subset(edited_subset: pd.DataFrame) -> None:
    if edited_subset is None:
        return
    working_df = _working_df().copy()
    for idx in edited_subset.index:
        for col in edited_subset.columns:
            if col in working_df.columns:
                working_df.at[idx, col] = edited_subset.at[idx, col]
    _set_working_df(working_df)



def _set_review_status(indices: Iterable[int], status: str) -> None:
    working_df = _working_df().copy()
    idxs = list(indices)
    if idxs:
        working_df.loc[idxs, "Review Status"] = status
        _set_working_df(working_df)



def _blank_fields(indices: Iterable[int], fields: list[str]) -> None:
    working_df = _working_df().copy()
    idxs = list(indices)
    for idx in idxs:
        for field in fields:
            if field in working_df.columns:
                working_df.at[idx, field] = ""
        working_df.at[idx, "Review Status"] = "Needs Review"
    _set_working_df(working_df)



def _append_pattern_note(indices: Iterable[int], note: str) -> None:
    working_df = _working_df().copy()
    idxs = list(indices)
    for idx in idxs:
        current = str(working_df.at[idx, "Pattern Notes"] or "").strip()
        if note.lower() in current.lower():
            continue
        working_df.at[idx, "Pattern Notes"] = f"{current}; {note}".strip("; ")
    _set_working_df(working_df)



def _apply_bulk_accept_95() -> int:
    working_df = _working_df().copy()
    mask = working_df["Confidence"].fillna(0.0).astype(float) >= AUTO_ACCEPT_THRESHOLD
    count = int(mask.sum())
    if count:
        working_df.loc[mask, "Review Status"] = "Auto-Accepted"
        _set_working_df(working_df)
    return count



def _apply_bulk_accept_safe_matches() -> int:
    working_df = _working_df().copy()
    mask = _safe_pattern_mask(working_df)
    count = int(mask.sum())
    if count:
        working_df.loc[mask, "Review Status"] = "Auto-Accepted"
        _set_working_df(working_df)
    return count



def _label_map(df: pd.DataFrame) -> dict[str, int]:
    return {
        f"{row['Code']} · {row['Brand'] or '(blank brand)'} · {row['Model'] or '(blank model)'} · p{row['Page']} · row {idx+1}": idx
        for idx, row in df.iterrows()
    }



def _save_current_supplier_memory(supplier_name: str, profile_notes: str, custom_aliases: dict[str, str], protected_phrases: list[str]) -> str | None:
    try:
        save_profile(
            supplier_name=supplier_name,
            notes=profile_notes,
            custom_aliases=custom_aliases,
            protected_phrases=protected_phrases,
        )
        return None
    except Exception as exc:  # noqa: BLE001
        return str(exc)



def _queue_indices(df: pd.DataFrame) -> list[int]:
    queued = df[df["Review Status"] != "Auto-Accepted"]
    return list(queued.index)



def _next_queue_index(df: pd.DataFrame, current_idx: int) -> int:
    queue = _queue_indices(df)
    if not queue:
        return current_idx
    if current_idx not in queue:
        return queue[0]
    pos = queue.index(current_idx)
    return queue[(pos + 1) % len(queue)]



def _status_metrics(df: pd.DataFrame) -> dict[str, int]:
    return {
        "needs_review": int((df["Review Status"] == "Needs Review").sum()),
        "review_suggested": int((df["Review Status"] == "Review Suggested").sum()),
        "auto_accepted": int((df["Review Status"] == "Auto-Accepted").sum()),
        "missing_brand": int((df["Brand"].fillna("") == "").sum()),
        "missing_model": int((df["Model"].fillna("") == "").sum()),
        "missing_year": int((df["Year"].fillna("") == "").sum()),
        "missing_engine": int((df["Engine"].fillna("") == "").sum()),
    }


st.title("Smart Automotive Pricelist PDF to Excel")
st.caption(
    "Accuracy-first conversion for automotive PDF catalogs: layout-aware extraction, segment splitting, protected phrases, field validators, supplier memory, conflict logs, review queues, and faster approval controls."
)

with st.sidebar:
    st.subheader("Supplier memory")
    profiles = list_profiles()
    st.caption(f"Saved profiles: {len(profiles)}")
    for profile in profiles[:8]:
        st.write(f"- **{profile['supplier_name']}** · updated {profile['updated_at']}")

    st.divider()
    st.subheader("Built-in pattern handling")
    st.markdown(
        "- `TOY` -> **TOYOTA**  \n"
        "- `MIT` / `MITS` -> **MITSUBISHI**  \n"
        "- `07 / 07' / '07 / \"07` -> **2007** when used as a year  \n"
        "- `FR / FRT / RR / LH / RH / LOW / UP / INNER / OUTER` -> normalized fields  \n"
        "- keeps protected model phrases like `BONGO 2000`, `CUBE 1.5`, and `MIRAGE G4` from being misread as year/engine"
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
profile = get_profile(supplier_name) if supplier_name.strip() else None
profile_notes_default = (profile or {}).get("notes", "")
custom_aliases_default = (profile or {}).get("custom_aliases", {})
protected_phrases_default = (profile or {}).get("protected_phrases", [])

setup_left, setup_center = st.columns([1.2, 1.8])
with setup_left:
    profile_notes = st.text_area(
        "Supplier notes",
        value=profile_notes_default,
        height=110,
        placeholder="Example: uses TOY / MIT / NIS abbreviations and often chains multiple applications with '/'.",
    )
with setup_center:
    with st.expander("Optional supplier memory controls"):
        alias_text = st.text_area(
            "Custom aliases (one per line, format: short = full brand)",
            value="\n".join(f"{k} = {v}" for k, v in custom_aliases_default.items()),
            height=110,
        )
        protected_text = st.text_area(
            "Protected phrases (one per line)",
            value="\n".join(protected_phrases_default),
            height=90,
            placeholder="Example: BONGO 2000",
        )

uploaded = st.file_uploader("Upload automotive pricelist PDF", type=["pdf"])
custom_aliases = _parse_alias_text(alias_text)
protected_phrases = [line.strip() for line in protected_text.splitlines() if line.strip()]

if uploaded is not None:
    file_bytes = uploaded.getvalue()
    sig = _signature(uploaded.name, file_bytes, supplier_name, profile_notes, custom_aliases, protected_phrases)
    with st.spinner("Running layout-aware extraction and normalization..."):
        parsed = _cached_parse_catalog(
            uploaded.name,
            file_bytes,
            tuple(sorted(custom_aliases.items())),
            tuple(sorted(protected_phrases)),
        )

    base_applications_df = applications_dataframe(parsed.application_rows)
    _ensure_working_state(base_applications_df, sig)

    raw_df = raw_rows_dataframe(parsed.raw_rows)
    conflict_df = conflict_dataframe(parsed.conflict_rows)
    evidence_df = evidence_dataframe(parsed.evidence_rows)
    working_df = _working_df()
    review_df = working_df[working_df["Review Status"] != "Auto-Accepted"].copy()

    metrics = _status_metrics(working_df)
    st.subheader("Conversion summary")
    metric_cols = st.columns(6)
    metric_cols[0].metric("Catalog rows", parsed.metrics.get("catalog_rows", len(raw_df)))
    metric_cols[1].metric("Application rows", parsed.metrics.get("application_rows", len(working_df)))
    metric_cols[2].metric("Auto-accepted", metrics["auto_accepted"])
    metric_cols[3].metric("Review suggested", metrics["review_suggested"])
    metric_cols[4].metric("Needs review", metrics["needs_review"])
    metric_cols[5].metric("Conflict log rows", parsed.metrics.get("conflicts", len(conflict_df)))

    if parsed.warnings:
        for warning in parsed.warnings:
            st.warning(warning)

    st.info(
        "The app now includes faster review controls: bulk acceptance for safe rows, conflict-only filters, blank-important-field filters, editable queues, and row-level actions before export."
    )

    tabs = st.tabs(
        [
            "1) Setup + extraction",
            "2) Split preview",
            "3) Review dashboard",
            "4) Row detail",
            "5) Raw catalog rows",
            "6) Conflict log",
            "7) Export",
        ]
    )

    with tabs[0]:
        left, center, right = st.columns([1.0, 1.4, 1.0])
        with left:
            st.markdown("**Supplier memory in use**")
            st.write(f"Supplier: **{supplier_name.strip() or 'Unnamed'}**")
            st.write(f"Custom aliases loaded: **{len(custom_aliases)}**")
            st.write(f"Protected phrases loaded: **{len(protected_phrases)}**")
            if st.button("Save current supplier rule/template", key="setup_save_memory"):
                err = _save_current_supplier_memory(supplier_name, profile_notes, custom_aliases, protected_phrases)
                if err:
                    st.error(err)
                else:
                    st.success(f"Saved supplier memory for {supplier_name.strip()}.")
        with center:
            st.markdown("**Extraction summary**")
            summary_df = pd.DataFrame(
                {
                    "Metric": ["Catalog rows", "Application rows", "Conflict log rows", "Evidence log rows"],
                    "Value": [
                        parsed.metrics.get("catalog_rows", 0),
                        parsed.metrics.get("application_rows", 0),
                        parsed.metrics.get("conflicts", 0),
                        parsed.metrics.get("evidence_rows", 0),
                    ],
                }
            )
            st.dataframe(summary_df, hide_index=True, use_container_width=True)
        with right:
            st.markdown("**Fast review actions available now**")
            st.markdown(
                "- Accept all 95%+ rows\n"
                "- Accept all safe rule matches\n"
                "- Show only conflicts\n"
                "- Show only blank important fields\n"
                "- Show only low confidence rows\n"
                "- Row-level accept / reject / blank / next-review controls"
            )

    with tabs[1]:
        st.subheader("Split preview")
        filter_cols = st.columns([1, 1, 1, 1, 1, 1])
        status_filter = filter_cols[0].selectbox(
            "Show rows",
            options=["All", "Auto-Accepted", "Review Suggested", "Needs Review"],
            index=0,
            key="split_status_filter",
        )
        code_filter = filter_cols[1].text_input("Filter by code", placeholder="Example: VKX-1267", key="split_code_filter")
        brand_filter = filter_cols[2].text_input("Filter by vehicle brand", placeholder="Example: TOYOTA", key="split_brand_filter")
        page_filter = filter_cols[3].text_input("Filter by page", placeholder="Example: 13", key="split_page_filter")
        conflict_only = filter_cols[4].checkbox("Show only conflicts", key="split_conflicts_only")
        blanks_only = filter_cols[5].checkbox("Show only blank important fields", key="split_blanks_only")

        preview_df = working_df.copy()
        if status_filter != "All":
            preview_df = preview_df[preview_df["Review Status"] == status_filter]
        if code_filter.strip():
            preview_df = preview_df[preview_df["Code"].fillna("").astype(str).str.contains(code_filter.strip(), case=False, na=False)]
        if brand_filter.strip():
            preview_df = preview_df[preview_df["Brand"].fillna("").astype(str).str.contains(brand_filter.strip(), case=False, na=False)]
        if page_filter.strip().isdigit():
            preview_df = preview_df[preview_df["Page"].astype(str) == page_filter.strip()]
        if conflict_only:
            keys = _conflict_key_set(conflict_df)
            preview_df = preview_df[preview_df.index.map(lambda i: _row_key_for_idx(working_df, i) in keys)]
        if blanks_only:
            preview_df = preview_df[_important_blank_mask(preview_df)]

        st.caption("Edit final export values here. Changes are saved in the current session and included in the downloaded Excel workbook.")
        edited_subset = st.data_editor(
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
            key="applications_editor_v7",
        )
        _merge_editor_subset(edited_subset)
        st.success(f"Currently tracking {len(_working_df())} application rows in the review session.")

    with tabs[2]:
        st.subheader("Review dashboard")
        left, center, right = st.columns([0.95, 1.55, 0.95])

        with left:
            st.markdown("**Bulk actions**")
            if st.button("Accept all 95%+ rows", use_container_width=True):
                count = _apply_bulk_accept_95()
                st.success(f"Auto-accepted {count} row(s).")
            if st.button("Accept all safe rule matches", use_container_width=True):
                count = _apply_bulk_accept_safe_matches()
                st.success(f"Auto-accepted {count} safe rule-match row(s).")
            st.divider()
            st.markdown("**Quick filters**")
            if st.button("Show only conflicts", use_container_width=True):
                st.session_state["dashboard_queue_filter"] = "Conflicts only"
            if st.button("Show only blank important fields", use_container_width=True):
                st.session_state["dashboard_queue_filter"] = "Blank important fields"
            if st.button(f"Show only low confidence (<{int(LOW_CONFIDENCE_THRESHOLD * 100)}%)", use_container_width=True):
                st.session_state["dashboard_queue_filter"] = f"Low confidence (<{int(LOW_CONFIDENCE_THRESHOLD * 100)}%)"
            if st.button("Reset dashboard filters", use_container_width=True):
                st.session_state["dashboard_queue_filter"] = "All review rows"
                st.session_state["dashboard_search_code"] = ""
                st.session_state["dashboard_search_brand"] = ""

            st.divider()
            queue_filter = st.selectbox(
                "Queue filter",
                options=_queue_filter_options(),
                key="dashboard_queue_filter",
            )
            st.text_input("Search code", key="dashboard_search_code", placeholder="Example: VKX-1267")
            st.text_input("Search brand", key="dashboard_search_brand", placeholder="Example: TOYOTA")

        with center:
            current_df = _working_df()
            queue_df = _apply_queue_filter(current_df, st.session_state["dashboard_queue_filter"], conflict_df)
            queue_df = _apply_search_filters(queue_df, st.session_state.get("dashboard_search_code", ""), st.session_state.get("dashboard_search_brand", ""))

            dash_cols = st.columns(6)
            dash_metrics = _status_metrics(current_df)
            dash_cols[0].metric("Needs review", dash_metrics["needs_review"])
            dash_cols[1].metric("Review suggested", dash_metrics["review_suggested"])
            dash_cols[2].metric("Missing brand", dash_metrics["missing_brand"])
            dash_cols[3].metric("Missing model", dash_metrics["missing_model"])
            dash_cols[4].metric("Missing year", dash_metrics["missing_year"])
            dash_cols[5].metric("Missing engine", dash_metrics["missing_engine"])

            label_map = _label_map(queue_df)
            selected_labels = st.multiselect(
                "Select rows for bulk row actions",
                options=list(label_map.keys()),
                key="dashboard_selected_labels",
                placeholder="Choose one or more rows from the current queue",
            )
            selected_indices = [label_map[label] for label in selected_labels]

            row_action_cols = st.columns(4)
            if row_action_cols[0].button("Accept selected rows", use_container_width=True):
                _set_review_status(selected_indices, "Auto-Accepted")
                st.success(f"Accepted {len(selected_indices)} selected row(s).")
            if row_action_cols[1].button("Mark selected for review", use_container_width=True):
                _set_review_status(selected_indices, "Needs Review")
                st.success(f"Moved {len(selected_indices)} selected row(s) to Needs Review.")
            if row_action_cols[2].button("Reject AI proposal", use_container_width=True):
                _set_review_status(selected_indices, "Needs Review")
                _append_pattern_note(selected_indices, "AI proposal rejected during review")
                st.success(f"Rejected AI proposal for {len(selected_indices)} selected row(s).")
            if row_action_cols[3].button("Save current supplier rule/template", use_container_width=True):
                err = _save_current_supplier_memory(supplier_name, profile_notes, custom_aliases, protected_phrases)
                if err:
                    st.error(err)
                else:
                    st.success(f"Saved supplier memory for {supplier_name.strip()}.")

            st.dataframe(queue_df, use_container_width=True, hide_index=True)

        with right:
            st.markdown("**What these buttons do**")
            st.markdown(
                "- **Accept all 95%+ rows**: clears easy approvals fast\n"
                "- **Accept all safe rule matches**: clears direct alias/year/token wins\n"
                "- **Show only conflicts**: focuses on rows with competing interpretations\n"
                "- **Show only blank important fields**: focuses on missing Brand/Model/Year/Engine\n"
                "- **Accept selected rows**: quickly clears rows you trust\n"
                "- **Reject AI proposal**: leaves the row in review and logs the rejection"
            )
            st.divider()
            st.markdown("**Review shortcuts**")
            st.write(f"Rows in current queue: **{len(queue_df)}**")
            st.write(f"Selected rows: **{len(st.session_state.get('dashboard_selected_labels', []))}**")
            st.write(f"Conflicts logged: **{len(conflict_df)}**")

    with tabs[3]:
        st.subheader("Row detail")
        current_df = _working_df()
        queue_idxs = _queue_indices(current_df)
        if not queue_idxs:
            st.success("All rows are currently auto-accepted. You can still review the Split Preview or export the workbook.")
        else:
            detail_label_map = _label_map(current_df.loc[queue_idxs])
            labels = list(detail_label_map.keys())
            default_pointer = min(st.session_state.get("detail_pointer", 0), len(labels) - 1)
            selected_label = st.selectbox("Inspect review row", options=labels, index=default_pointer)
            selected_idx = detail_label_map[selected_label]
            st.session_state["detail_pointer"] = labels.index(selected_label)
            selected_row = current_df.loc[selected_idx].copy()
            row_key = _row_key_for_idx(current_df, selected_idx)

            left, center, right = st.columns([1.0, 1.25, 1.0])
            with left:
                st.markdown("**Raw source context**")
                st.code(str(selected_row["Source Line"]))
                st.markdown("**Current parsed fields**")
                st.write(selected_row[["Code", "Axle", "Side", "Vertical", "Mount", "Brand", "Model", "Engine", "Year", "Review Status"]])
                st.markdown("**Field selection for row actions**")
                selected_fields = st.multiselect(
                    "Selected fields",
                    options=EDITABLE_FIELDS,
                    default=["Brand", "Model", "Year", "Engine"],
                    key=f"detail_selected_fields_{selected_idx}",
                )

            with center:
                st.markdown("**Edit manually**")
                with st.form(f"row_edit_form_{selected_idx}"):
                    field_inputs: dict[str, object] = {}
                    input_cols = st.columns(2)
                    for i, field in enumerate(EDITABLE_FIELDS):
                        target_col = input_cols[i % 2]
                        current_value = selected_row.get(field, "")
                        with target_col:
                            if field in {"Your Price (PHP)", "Use Price (PHP)"}:
                                current_num = None if pd.isna(current_value) or current_value == "" else float(current_value)
                                field_inputs[field] = st.number_input(field, value=current_num or 0.0, step=1.0, format="%.2f")
                            else:
                                field_inputs[field] = st.text_input(field, value="" if pd.isna(current_value) else str(current_value))
                    save_row = st.form_submit_button("Save row edits", use_container_width=True)
                if save_row:
                    updated_df = _working_df().copy()
                    for field, value in field_inputs.items():
                        updated_df.at[selected_idx, field] = value
                    _set_working_df(updated_df)
                    st.success("Saved manual edits for the selected row.")

                action_cols = st.columns(3)
                if action_cols[0].button("Accept row", use_container_width=True, key=f"accept_row_{selected_idx}"):
                    _set_review_status([selected_idx], "Auto-Accepted")
                    st.success("Accepted this row.")
                if action_cols[1].button("Accept selected fields", use_container_width=True, key=f"accept_fields_{selected_idx}"):
                    _append_pattern_note([selected_idx], f"reviewer accepted fields: {', '.join(selected_fields) if selected_fields else 'none'}")
                    if selected_row["Confidence"] >= AUTO_ACCEPT_THRESHOLD and not _important_blank_mask(current_df.loc[[selected_idx]]).iloc[0]:
                        _set_review_status([selected_idx], "Auto-Accepted")
                    else:
                        _set_review_status([selected_idx], "Review Suggested")
                    st.success("Accepted the selected fields for this row.")
                if action_cols[2].button("Reject AI proposal", use_container_width=True, key=f"reject_row_{selected_idx}"):
                    _set_review_status([selected_idx], "Needs Review")
                    _append_pattern_note([selected_idx], "AI proposal rejected during row review")
                    st.warning("Marked this row as Needs Review.")

                action_cols2 = st.columns(3)
                if action_cols2[0].button("Keep selected fields blank", use_container_width=True, key=f"blank_fields_{selected_idx}"):
                    _blank_fields([selected_idx], selected_fields)
                    st.warning("Blanked the selected fields and kept the row in review.")
                if action_cols2[1].button("Next review row", use_container_width=True, key=f"next_row_{selected_idx}"):
                    next_idx = _next_queue_index(_working_df(), selected_idx)
                    next_label_map = _label_map(_working_df().loc[_queue_indices(_working_df())])
                    next_labels = list(next_label_map.keys())
                    next_target_label = next((lab for lab, idx in next_label_map.items() if idx == next_idx), None)
                    if next_target_label in next_labels:
                        st.session_state["detail_pointer"] = next_labels.index(next_target_label)
                    st.rerun()
                if action_cols2[2].button("Save current supplier rule/template", use_container_width=True, key=f"detail_save_memory_{selected_idx}"):
                    phrase = f"{selected_row['Brand']} {selected_row['Model']}".strip()
                    if phrase and phrase not in protected_phrases:
                        protected_phrases.append(phrase)
                    err = _save_current_supplier_memory(supplier_name, profile_notes, custom_aliases, protected_phrases)
                    if err:
                        st.error(err)
                    else:
                        st.success(f"Saved supplier memory and protected phrase hint for: {phrase}")

            with right:
                st.markdown("**Evidence panel**")
                row_evidence = evidence_df[evidence_df["Row Key"] == row_key].copy()
                row_conflicts = conflict_df[conflict_df["Row Key"] == row_key].copy()
                st.write(f"Confidence: **{int(float(selected_row['Confidence']) * 100)}%**")
                st.write(f"Pattern notes: {selected_row['Pattern Notes'] or '—'}")
                st.write(f"Review status: **{selected_row['Review Status']}**")
                if not row_evidence.empty:
                    st.markdown("**Field evidence**")
                    st.dataframe(row_evidence, use_container_width=True, hide_index=True)
                if not row_conflicts.empty:
                    st.markdown("**Conflicts / cautions**")
                    st.dataframe(row_conflicts, use_container_width=True, hide_index=True)
                else:
                    st.success("No conflict rows logged for this record.")

    with tabs[4]:
        st.subheader("Raw catalog rows")
        st.caption("This is the raw extraction layer before splitting and normalization.")
        st.dataframe(raw_df, use_container_width=True, hide_index=True)

    with tabs[5]:
        st.subheader("Conflict log")
        st.caption("Use this sheet to inspect ambiguous interpretations, field-domain warnings, or inferred values.")
        st.dataframe(conflict_df, use_container_width=True, hide_index=True)

    with tabs[6]:
        st.subheader("Export")
        final_applications_df = _working_df().copy()
        review_export_df = final_applications_df[final_applications_df["Review Status"] != "Auto-Accepted"].copy()

        export_cols = st.columns([1, 1, 1])
        if export_cols[0].button("Save current supplier rule/template", use_container_width=True):
            err = _save_current_supplier_memory(supplier_name, profile_notes, custom_aliases, protected_phrases)
            if err:
                st.error(err)
            else:
                st.success(f"Saved supplier memory for {supplier_name.strip()}.")
        if export_cols[1].button("Accept all 95%+ before export", use_container_width=True):
            count = _apply_bulk_accept_95()
            final_applications_df = _working_df().copy()
            review_export_df = final_applications_df[final_applications_df["Review Status"] != "Auto-Accepted"].copy()
            st.success(f"Auto-accepted {count} row(s) before export.")
        if export_cols[2].button("Show only blank important fields in dashboard", use_container_width=True):
            st.session_state["dashboard_queue_filter"] = "Blank important fields"
            st.info("Dashboard filter set to Blank important fields.")

        workbook = build_workbook(final_applications_df, raw_df, review_export_df, conflict_df, evidence_df)
        excel_bytes = workbook_to_bytes(workbook)
        output_name = f"{Path(uploaded.name).stem}_review_buttons.xlsx"
        st.download_button(
            "Download Excel workbook",
            data=excel_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        st.markdown("**Workbook sheets**")
        st.markdown(
            "- **Applications**: final normalized rows after your edits and review decisions  \n"
            "- **Catalog Rows**: raw extracted line-level rows  \n"
            "- **Review Queue**: rows not auto-accepted  \n"
            "- **Conflict Log**: ambiguity, validation, or inference warnings  \n"
            "- **Evidence Log**: field-by-field reasons and confidence"
        )
