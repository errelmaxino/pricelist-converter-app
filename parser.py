from __future__ import annotations

import datetime as _dt
import hashlib
import io
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

CANONICAL_COLUMNS = [
    "Category",
    "Brand",
    "Part No.",
    "OE No.",
    "Model",
    "Year",
    "Size",
    "Original Price (PHP)",
    "Your Price (PHP)",
    "Use Price (PHP)",
    "Page",
    "Source Line",
]

KEEP_AS_EXTRA = "__KEEP_AS_EXTRA__"
IGNORE_COLUMN = "__IGNORE__"

HEADER_ALIASES = {
    "Category": ["category", "product category", "group", "section"],
    "Brand": ["brand", "make", "manufacturer", "supplier brand"],
    "Part No.": [
        "part no",
        "part number",
        "part #",
        "item code",
        "item no",
        "sku",
        "code",
        "product code",
    ],
    "OE No.": ["oe no", "oe number", "oem no", "oem number", "oe #", "oem #", "reference no"],
    "Model": ["model", "application", "vehicle", "fitment", "description", "item description"],
    "Year": ["year", "yr", "year model", "model year", "years"],
    "Size": ["size", "dia", "diameter", "spec", "specification", "dimension"],
    "Original Price (PHP)": [
        "price",
        "amount",
        "list price",
        "srp",
        "unit price",
        "original price",
        "selling price",
        "php",
        "price php",
    ],
}

PRICE_RE = re.compile(r"(?:PHP|₱)?\s*\d{1,3}(?:,\d{3})*(?:\.\d{2})?")
HEADER_WORD_RE = re.compile(r"^[A-Za-z][A-Za-z0-9#()./\-\s]*$")


def _clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ").replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def canonicalize_header(header: str) -> Optional[str]:
    normalized = _slug(header)
    if not normalized:
        return None
    for canonical, aliases in HEADER_ALIASES.items():
        if normalized == _slug(canonical):
            return canonical
        for alias in aliases:
            if normalized == _slug(alias):
                return canonical
    return None


@dataclass
class ParsedDocument:
    detected_headers: List[str]
    rows: List[Dict[str, object]]
    header_fingerprint: str
    warnings: List[str] = field(default_factory=list)


# ----- year / model cleanup rules -----

def expand_short_year(two_digit: str) -> int:
    year = int(two_digit)
    pivot = (_dt.date.today().year % 100) + 1
    return 2000 + year if year <= pivot else 1900 + year


def normalize_year_expression(raw: str) -> str:
    value = re.sub(r"\s+", " ", raw.strip().replace("’", "'")).replace("-up", "-Up")

    full = re.fullmatch(r"((?:19|20)\d{2})\s*-\s*((?:19|20)\d{2}|[Uu]p)", value)
    if full:
        end = "Up" if full.group(2).lower() == "up" else full.group(2)
        return f"{full.group(1)} - {end}"

    short = re.fullmatch(r"'(\d{2})\s*-\s*'?(?:(\d{2})|([Uu]p))?", value)
    if short:
        start = expand_short_year(short.group(1))
        if short.group(2):
            return f"{start} - {expand_short_year(short.group(2))}"
        return f"{start} - Up"

    short_up = re.fullmatch(r"'(\d{2})\s*([Uu]p)", value)
    if short_up:
        return f"{expand_short_year(short_up.group(1))} - Up"

    single = re.fullmatch(r"((?:19|20)\d{2})", value)
    if single:
        return single.group(1)

    return value


def split_model_and_year(model_year_text: str) -> Tuple[str, str]:
    text = re.sub(r"\s+", " ", model_year_text.strip().replace("’", "'"))
    if not text:
        return "", ""

    patterns = [
        re.compile(
            r"^(?P<model>.*?)(?P<year>(?:19|20)\d{2}\s*-\s*(?:19|20)\d{2}|(?:19|20)\d{2}\s*-\s*[Uu]p|(?:19|20)\d{2})$"
        ),
        re.compile(r"^(?P<model>.*?)(?P<year>'\d{2}\s*-\s*'?(?:\d{2})|'\d{2}\s*-\s*[Uu]p|'\d{2}\s*-\s*|'\d{2}\s*[Uu]p)$"),
    ]
    for pattern in patterns:
        match = pattern.match(text)
        if match:
            model = match.group("model").strip(" -")
            year = normalize_year_expression(match.group("year"))
            return model, year
    return text, ""


def fix_cross_column_year(model: str, year: str, size_text: str) -> Tuple[str, str, str]:
    model = re.sub(r"\s+", " ", model).strip()
    year = re.sub(r"\s+", " ", year).strip()
    size_text = re.sub(r"\s+", " ", size_text).strip().replace("’", "'")

    if not year:
        match = re.match(r"^(?P<model>.*?)(?P<start>(?:19|20)\d{2})\s*-\s*$", model)
        size_match = re.match(r"^(?P<end>(?:19|20)\d{2})\s+(?P<size>.+)$", size_text)
        if match and size_match:
            return match.group("model").strip(" -"), f"{match.group('start')} - {size_match.group('end')}", size_match.group("size").strip()

    if not year:
        shorthand = re.search(
            r"^(?P<model>.*?)(?:\s+)('(?:\d{2})\s*-\s*'?(?:\d{2})|'\d{2}\s*-\s*[Uu]p|'\d{2}\s*-\s*|'\d{2}\s*[Uu]p)$",
            model,
        )
        if shorthand:
            model = shorthand.group("model").strip()
            year = normalize_year_expression(shorthand.group(2))

    if not year:
        size_year = re.match(r"^(?P<start>(?:19|20)\d{2})\s*-\s*(?P<end>(?:19|20)\d{2}|[Uu]p)\s+(?P<size>.+)$", size_text)
        if size_year:
            end = "Up" if size_year.group("end").lower() == "up" else size_year.group("end")
            return model, f"{size_year.group('start')} - {end}", size_year.group("size").strip()

    return model, year, size_text


# ----- pdf extraction -----

def _group_lines(words: List[dict], tolerance: float = 2.0) -> List[dict]:
    lines: List[dict] = []
    for word in sorted(words, key=lambda x: (x["top"], x["x0"])):
        placed = False
        for line in lines:
            if abs(line["top"] - word["top"]) <= tolerance:
                line["words"].append(word)
                placed = True
                break
        if not placed:
            lines.append({"top": word["top"], "words": [word]})

    output = []
    for line in sorted(lines, key=lambda x: x["top"]):
        row_words = sorted(line["words"], key=lambda x: x["x0"])
        output.append(
            {
                "top": line["top"],
                "words": row_words,
                "text": " ".join(w["text"] for w in row_words).strip(),
                "min_x": min(w["x0"] for w in row_words),
            }
        )
    return output


def _is_heading_line(text: str) -> bool:
    text = text.strip()
    upper = text.upper()
    if not text or re.search(r"\d", text):
        return False
    if PRICE_RE.search(text):
        return False
    if any(token in upper for token in ["PART NUMBER", "OE NUMBER", "MODEL", "YEAR", "SIZE", "PRICE"]):
        return False
    letters_only = re.sub(r"[^A-Za-z&/\-\s]", "", text)
    return bool(letters_only) and letters_only.upper() == letters_only


def _clean_table(table: List[List[object]]) -> List[List[str]]:
    cleaned: List[List[str]] = []
    max_len = max((len(r) for r in table), default=0)
    for row in table:
        values = [_clean_text(cell) for cell in row]
        if len(values) < max_len:
            values.extend([""] * (max_len - len(values)))
        if any(values):
            cleaned.append(values)
    return cleaned


def _score_header_row(row: List[str]) -> int:
    non_empty = [c for c in row if c]
    if len(non_empty) < 2:
        return -1
    canonical_hits = sum(1 for cell in non_empty if canonicalize_header(cell))
    alpha_hits = sum(1 for cell in non_empty if HEADER_WORD_RE.match(cell) and not re.search(r"\d", cell))
    if canonical_hits >= 2:
        return 10 + canonical_hits
    if canonical_hits >= 1 and alpha_hits >= 2:
        return 7 + canonical_hits
    if alpha_hits >= max(3, len(non_empty) - 1):
        return 4
    return 0


def _looks_like_data_row(row: List[str]) -> bool:
    text = " | ".join(row)
    return bool(re.search(r"\d", text) or PRICE_RE.search(text))


def _dedupe_headers(headers: List[str]) -> List[str]:
    seen: Dict[str, int] = {}
    output: List[str] = []
    for idx, header in enumerate(headers, start=1):
        clean = _clean_text(header) or f"Column {idx}"
        count = seen.get(clean, 0) + 1
        seen[clean] = count
        output.append(clean if count == 1 else f"{clean} ({count})")
    return output


def _detect_headings_for_table(page, table_top: float, current_category: str, current_brand: str) -> Tuple[str, str]:
    words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
    lines = _group_lines(words)
    for line in lines:
        if line["top"] >= table_top:
            break
        text = line["text"]
        if _is_heading_line(text) and line["min_x"] < 140:
            if len(text.split()) > 1:
                current_category = text.title()
            else:
                current_brand = text.title()
    return current_category, current_brand


def extract_document(pdf_path: str | Path) -> ParsedDocument:
    rows: List[Dict[str, object]] = []
    all_headers: List[str] = []
    warnings: List[str] = []
    last_headers: Optional[List[str]] = None
    current_category = ""
    current_brand = ""

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            tables = page.find_tables()
            if not tables:
                continue

            for table in tables:
                current_category, current_brand = _detect_headings_for_table(page, table.bbox[1], current_category, current_brand)
                raw_rows = _clean_table(table.extract())
                if not raw_rows:
                    continue

                header_score = _score_header_row(raw_rows[0])
                if header_score >= 7:
                    headers = _dedupe_headers(raw_rows[0])
                    data_rows = raw_rows[1:]
                    last_headers = headers
                elif last_headers and len(last_headers) == len(raw_rows[0]) and _looks_like_data_row(raw_rows[0]):
                    headers = last_headers
                    data_rows = raw_rows
                else:
                    headers = _dedupe_headers([f"Column {i}" for i in range(1, len(raw_rows[0]) + 1)])
                    data_rows = raw_rows
                    warnings.append(f"Page {page_number}: used generic column names for one detected table.")

                for h in headers:
                    if h not in all_headers:
                        all_headers.append(h)

                for raw in data_rows:
                    if len(raw) < len(headers):
                        raw = raw + [""] * (len(headers) - len(raw))
                    elif len(raw) > len(headers):
                        raw = raw[: len(headers)]

                    row = {headers[i]: raw[i] for i in range(len(headers))}
                    row["Category"] = current_category
                    row["Brand"] = current_brand
                    row["Page"] = page_number
                    row["Source Line"] = " | ".join(v for v in raw if v)
                    rows.append(row)

    if not rows:
        warnings.append("No tables were detected. This app works best with text-based tabular PDFs.")

    base_headers = list(all_headers)
    if any(r.get("Category") for r in rows) and "Category" not in base_headers:
        base_headers.insert(0, "Category")
    if any(r.get("Brand") for r in rows) and "Brand" not in base_headers:
        insert_at = 1 if "Category" in base_headers else 0
        base_headers.insert(insert_at, "Brand")

    fingerprint_source = "|".join(_slug(h) for h in base_headers)
    fingerprint = hashlib.md5(fingerprint_source.encode("utf-8")).hexdigest()[:12] if fingerprint_source else "noheaders"
    return ParsedDocument(detected_headers=base_headers, rows=rows, header_fingerprint=fingerprint, warnings=warnings)


# ----- mapping / output -----

def build_suggested_mapping(headers: Iterable[str]) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    used_targets = set()
    for header in headers:
        target = canonicalize_header(header)
        if target and target not in used_targets:
            mapping[header] = target
            used_targets.add(target)
        elif header in {"Category", "Brand"} and header not in used_targets:
            mapping[header] = header
            used_targets.add(header)
        else:
            mapping[header] = KEEP_AS_EXTRA
    return mapping


def apply_mapping(
    doc: ParsedDocument,
    mapping: Dict[str, str],
    mode: str = "normalized",
    preserve_unknown: bool = True,
) -> Tuple[List[str], List[Dict[str, object]]]:
    if mode == "exact":
        exact_headers = list(doc.detected_headers)
        output_headers = exact_headers + [h for h in ["Page", "Source Line"] if h not in exact_headers]
        out_rows: List[Dict[str, object]] = []
        for raw in doc.rows:
            row = {h: raw.get(h, "") for h in output_headers}
            out_rows.append(row)
        return output_headers, out_rows

    extras_in_order: List[str] = []
    canonical_headers = list(CANONICAL_COLUMNS)
    out_rows: List[Dict[str, object]] = []

    for raw in doc.rows:
        normalized = {col: "" for col in canonical_headers}
        extras: Dict[str, object] = {}

        for source_header in doc.detected_headers:
            if source_header in {"Page", "Source Line"}:
                continue
            value = raw.get(source_header, "")
            target = mapping.get(source_header, KEEP_AS_EXTRA)
            if target in CANONICAL_COLUMNS:
                if not normalized[target]:
                    normalized[target] = value
                elif value and str(value) not in str(normalized[target]):
                    normalized[target] = f"{normalized[target]} | {value}"
            elif target == KEEP_AS_EXTRA and preserve_unknown:
                extras[source_header] = value
                if source_header not in extras_in_order:
                    extras_in_order.append(source_header)

        # carry metadata if missing from mapping
        normalized["Category"] = normalized["Category"] or raw.get("Category", "")
        normalized["Brand"] = normalized["Brand"] or raw.get("Brand", "")
        normalized["Page"] = raw.get("Page", "")
        normalized["Source Line"] = raw.get("Source Line", "")

        model, year = split_model_and_year(_clean_text(normalized["Model"]))
        merged_model = model or _clean_text(normalized["Model"])
        merged_year = year or _clean_text(normalized["Year"])
        fixed_model, fixed_year, fixed_size = fix_cross_column_year(merged_model, merged_year, _clean_text(normalized["Size"]))
        normalized["Model"] = fixed_model
        normalized["Year"] = fixed_year
        normalized["Size"] = fixed_size

        price_value = _coerce_price(normalized["Original Price (PHP)"])
        normalized["Original Price (PHP)"] = price_value if price_value is not None else _clean_text(normalized["Original Price (PHP)"])
        normalized["Your Price (PHP)"] = None
        normalized["Use Price (PHP)"] = None

        for extra_header, extra_value in extras.items():
            normalized[extra_header] = extra_value
        out_rows.append(normalized)

    output_headers = canonical_headers.copy()
    if preserve_unknown:
        insert_at = output_headers.index("Page")
        for extra in extras_in_order:
            if extra not in output_headers:
                output_headers.insert(insert_at, extra)
                insert_at += 1
    return output_headers, out_rows


def _coerce_price(value: object) -> Optional[float]:
    text = _clean_text(value)
    if not text:
        return None
    match = PRICE_RE.search(text)
    if not match:
        return None
    amount = re.sub(r"[^\d.]", "", match.group(0))
    try:
        return float(amount)
    except ValueError:
        return None


def dataframe_preview(headers: List[str], rows: List[Dict[str, object]], limit: int = 25):
    import pandas as pd

    preview_rows = []
    for row in rows[:limit]:
        preview_rows.append({h: row.get(h, "") for h in headers})
    return pd.DataFrame(preview_rows, columns=headers)


# ----- excel -----

def build_workbook(headers: List[str], rows: List[Dict[str, object]]) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Pricelist"

    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    your_price_col = headers.index("Your Price (PHP)") + 1 if "Your Price (PHP)" in headers else None
    use_price_col = headers.index("Use Price (PHP)") + 1 if "Use Price (PHP)" in headers else None
    original_price_col = headers.index("Original Price (PHP)") + 1 if "Original Price (PHP)" in headers else None

    for row_idx, row in enumerate(rows, start=2):
        for col_idx, header in enumerate(headers, start=1):
            value = row.get(header, "")
            if header == "Use Price (PHP)" and your_price_col and original_price_col:
                your_ref = f"{get_column_letter(your_price_col)}{row_idx}"
                orig_ref = f"{get_column_letter(original_price_col)}{row_idx}"
                ws.cell(row=row_idx, column=col_idx, value=f'=IF({your_ref}="",{orig_ref},{your_ref})')
            else:
                ws.cell(row=row_idx, column=col_idx, value=value)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    price_headers = {"Original Price (PHP)", "Your Price (PHP)", "Use Price (PHP)"}
    for col_idx, header in enumerate(headers, start=1):
        if header in price_headers:
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row=row_idx, column=col_idx).number_format = '₱#,##0.00'

    for idx, header in enumerate(headers, start=1):
        max_len = max(len(_clean_text(header)), *(len(_clean_text(ws.cell(r, idx).value)) for r in range(2, min(ws.max_row, 250) + 1)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 10), 40)

    return wb


def workbook_to_bytes(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def convert_pdf_to_excel_bytes(
    pdf_path: str | Path,
    mapping: Optional[Dict[str, str]] = None,
    mode: str = "normalized",
    preserve_unknown: bool = True,
) -> Tuple[bytes, ParsedDocument, List[str], List[Dict[str, object]]]:
    doc = extract_document(pdf_path)
    final_mapping = mapping or build_suggested_mapping(doc.detected_headers)
    headers, rows = apply_mapping(doc, final_mapping, mode=mode, preserve_unknown=preserve_unknown)
    wb = build_workbook(headers, rows)
    return workbook_to_bytes(wb), doc, headers, rows
