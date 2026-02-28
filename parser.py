from __future__ import annotations

import io
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# -----------------------------
# Static rules / normalizations
# -----------------------------

BRAND_ALIASES: Dict[str, str] = {
    "TOY": "TOYOTA",
    "TOY.": "TOYOTA",
    "TOYO": "TOYOTA",
    "TOYOYA": "TOYOTA",
    "TOYOTA": "TOYOTA",
    "MIT": "MITSUBISHI",
    "MIT.": "MITSUBISHI",
    "MITS": "MITSUBISHI",
    "MITS.": "MITSUBISHI",
    "MITSUBISHI": "MITSUBISHI",
    "MITSUBUSHI": "MITSUBISHI",
    "NIS": "NISSAN",
    "NIS.": "NISSAN",
    "NISSAN": "NISSAN",
    "ISU": "ISUZU",
    "ISU.": "ISUZU",
    "ISUZU": "ISUZU",
    "HYU": "HYUNDAI",
    "HYU.": "HYUNDAI",
    "HYUNDAI": "HYUNDAI",
    "MAZ": "MAZDA",
    "MAZ.": "MAZDA",
    "MAZDA": "MAZDA",
    "SUZ": "SUZUKI",
    "SUZ.": "SUZUKI",
    "SUZUKI": "SUZUKI",
    "KIA": "KIA",
    "HONDA": "HONDA",
    "FORD": "FORD",
    "CHEV": "CHEVROLET",
    "CHEV.": "CHEVROLET",
    "CHEVROLET": "CHEVROLET",
    "DAIHATSU": "DAIHATSU",
    "MERCEDES": "MERCEDES BENZ",
    "BENZ": "MERCEDES BENZ",
    "MERCEDES BENZ": "MERCEDES BENZ",
    "SUBARU": "SUBARU",
    "PROTON": "PROTON",
    "FOTON": "FOTON",
    "DAEWOO": "DAEWOO",
    "MAHINDRA": "MAHINDRA",
    "ASIATOPIC": "ASIATOPIC",
    "NUVO-PRO": "NUVO-PRO",
}

ENGINE_CODES = {
    "18R",
    "2E",
    "3K",
    "4K",
    "12R",
    "2C",
    "2L",
    "3L",
    "4D31",
    "4D32",
    "4DR5",
    "4HF1",
    "4BC2",
    "4M40",
    "4G63",
    "C240",
    "C-240",
    "4BA1",
    "4BB1",
    "4BC1",
    "4BD1",
    "4BE1",
    "4JA1",
    "4JB1",
    "4JG2",
    "4D56",
    "1KD",
    "2KD",
    "K6A",
}

BRAND_REGEX = re.compile(
    r"^(?P<brand>"
    + "|".join(re.escape(k) for k in sorted(BRAND_ALIASES, key=len, reverse=True))
    + r")\b[\s.-]*(?P<rest>.*)$",
    re.IGNORECASE,
)
ENGINE_REGEX = re.compile(
    r"\b(?:" + "|".join(re.escape(code) for code in sorted(ENGINE_CODES, key=len, reverse=True)) + r")\b",
    re.IGNORECASE,
)
CODE_REGEX = re.compile(r"^(?P<code>(?:VAX|VA|VKX|VK)-[A-Z0-9*]+)\s*(?P<rest>.*)$", re.IGNORECASE)
PRICE_REGEX = re.compile(r"(?P<price>\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\s*$")
AXLE_REGEX = re.compile(
    r"^(?P<axle>FRT\.?\s*&\s*RR|FRT\.?&\s*RR|FRT\s*&\s*RR|FRT-RR|FRT/?RR|FRONT|REAR|RR|FR|FRT\.?)\b[\s:.-]*(?P<rest>.*)$",
    re.IGNORECASE,
)

NOISE_TERMS = [
    "TIHLUCK",
    "PASAY",
    "EMAIL",
    "UPDATE SEPT",
    "OF 18",
    "(02)",
    "GUARANTEED",
    "TRADING",
    "CORPORATION",
    "GMAIL.COM",
]

MAIN_COLUMNS = [
    "Source Line",
    "Supplier Brand",
    "Category",
    "Code",
    "Axle",
    "Brand",
    "Model",
    "Engine",
    "Year",
    "Original Price (PHP)",
    "Your Price (PHP)",
    "Use Price (PHP)",
    "Page",
    "Confidence",
    "Review Status",
    "Pattern Notes",
]

RAW_COLUMNS = ["Page", "Supplier Brand", "Category", "Code", "Axle", "Description", "Original Price (PHP)"]

PROTECTED_MODEL_YEAR_PATTERNS = [
    re.compile(r"\bBONGO\s+2000\b", re.IGNORECASE),
]


@dataclass
class CatalogRow:
    page: int
    supplier_brand: str
    category: str
    code: str
    axle: str
    description: str
    price: Optional[float]
    source_line: str


@dataclass
class ParsedCatalog:
    raw_rows: List[CatalogRow]
    application_rows: List[Dict[str, object]]
    review_rows: List[Dict[str, object]]
    warnings: List[str]


# -----------------------------
# Text helpers
# -----------------------------


def _clean_text(text: str) -> str:
    value = str(text or "")
    value = (
        value.replace("–", "-")
        .replace("—", "-")
        .replace("’", "'")
        .replace("“", '"')
        .replace("”", '"')
        .replace("½", "1/2")
    )
    value = re.sub(r"(?<=\d)\s*,\s*(?=\d{3}\b)", ",", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def _cleanup_model(text: str) -> str:
    value = _clean_text(text)
    value = value.replace(".ORIG", " ORIG")
    value = value.replace("HI ACE", "HIACE")
    value = value.replace("STA FE", "SANTA FE")
    value = value.replace("I1O", "I10")
    value = re.sub(r"\b'\b", " ", value)
    value = value.replace("' ", " ").replace(" '", " ")
    value = value.replace('" ', ' ').replace(' "', ' ')
    value = value.replace('"', " ")
    value = value.replace("'", " ")
    value = re.sub(r"\s+", " ", value).strip(" ,-/")
    return value


def _model_key(model: str) -> str:
    return re.sub(r"[^A-Z0-9]+", " ", _cleanup_model(model).upper()).strip()


def _expand_short_year(two_digit: str) -> int:
    year = int(two_digit)
    return 2000 + year if year <= 35 else 1900 + year


def _normalize_year(value: str) -> str:
    text = _clean_text(value).replace('"', "'").upper().replace("MY", "")
    text = text.replace(" TO ", "-").replace("- UP", "-UP").replace(" - UP", "-UP")

    match = re.fullmatch(r"'?(\d{2})\s*-\s*'?(\d{2})", text)
    if match:
        return f"{_expand_short_year(match.group(1))}-{_expand_short_year(match.group(2))}"

    match = re.fullmatch(r"'?(\d{2})\s*-\s*UP", text)
    if match:
        return f"{_expand_short_year(match.group(1))}-Up"

    match = re.fullmatch(r"(\d{4})\s*-\s*(\d{4})", text)
    if match:
        return f"{match.group(1)}-{match.group(2)}"

    match = re.fullmatch(r"(\d{4})\s*-\s*UP", text)
    if match:
        return f"{match.group(1)}-Up"

    match = re.fullmatch(r"'?(\d{2})", text)
    if match:
        return str(_expand_short_year(match.group(1)))

    match = re.fullmatch(r"(\d{4})", text)
    if match:
        return match.group(1)

    return value


YEAR_PATTERNS = [
    re.compile(r"(?P<all>(?P<s>\d{4})\s*(?:-|TO)\s*(?P<e>\d{4}|UP))", re.IGNORECASE),
    # Short-year ranges like 93'-98', '93-'98, 93'-UP, '06-'20.
    re.compile(r"(?P<all>[\"']?(?P<s>\d{2})[\"']?\s*-\s*[\"']?(?P<e>\d{2}|UP)[\"']?)", re.IGNORECASE),
    re.compile(r"(?P<all>(?P<y>\d{4}))", re.IGNORECASE),
    re.compile(r"(?P<all>[\"']?(?P<y2>\d{2})[\"']?MY)", re.IGNORECASE),
    re.compile(r"(?P<all>[\"']?(?P<y3>\d{2})[\"'])", re.IGNORECASE),
    re.compile(r"(?P<all>(?<![A-Z0-9])(?P<y4>\d{2})(?![A-Z0-9]))$", re.IGNORECASE),
]


def _extract_year(text: str) -> Tuple[str, str]:
    protected_spans: List[Tuple[int, int]] = []
    for pattern in PROTECTED_MODEL_YEAR_PATTERNS:
        for protected in pattern.finditer(text):
            protected_spans.append(protected.span())

    best_match: Optional[Tuple[Tuple[int, int], re.Match[str]]] = None
    for pattern in YEAR_PATTERNS:
        for match in pattern.finditer(text):
            span = match.span("all")
            if any(span[0] >= p0 and span[1] <= p1 for p0, p1 in protected_spans):
                continue
            score = (span[1] - span[0], span[0])
            if best_match is None or score > best_match[0]:
                best_match = (score, match)

    if best_match is None:
        return "", text

    match = best_match[1]
    groups = match.groupdict()

    if groups.get("s") and groups.get("e"):
        start = groups["s"] if len(groups["s"]) == 4 else str(_expand_short_year(groups["s"]))
        end_group = groups["e"].upper()
        end = "Up" if end_group == "UP" else (end_group if len(end_group) == 4 else str(_expand_short_year(end_group)))
        year = f"{start}-{end}"
    elif groups.get("y"):
        year = groups["y"]
    elif groups.get("y2"):
        year = str(_expand_short_year(groups["y2"]))
    elif groups.get("y3"):
        year = str(_expand_short_year(groups["y3"]))
    elif groups.get("y4"):
        year = str(_expand_short_year(groups["y4"]))
    else:
        bare_two_digit = re.sub(r"[^0-9]", "", match.group("all"))
        year = str(_expand_short_year(bare_two_digit)) if len(bare_two_digit) == 2 else match.group("all")

    remaining = _cleanup_model((text[: match.start("all")] + " " + text[match.end("all") :]).strip())
    return year, remaining


# -----------------------------
# PDF reading / row extraction
# -----------------------------


def _group_page_lines(page) -> List[Dict[str, object]]:
    words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
    grouped: List[Dict[str, object]] = []
    for word in sorted(words, key=lambda x: (x["top"], x["x0"])):
        placed = False
        for line in grouped:
            if abs(float(line["top"]) - float(word["top"])) <= 2.2:
                line["words"].append(word)
                placed = True
                break
        if not placed:
            grouped.append({"top": word["top"], "words": [word]})

    output: List[Dict[str, object]] = []
    for line in sorted(grouped, key=lambda x: x["top"]):
        row_words = sorted(line["words"], key=lambda x: x["x0"])
        output.append(
            {
                "top": line["top"],
                "min_x": min(word["x0"] for word in row_words),
                "text": _clean_text(" ".join(word["text"] for word in row_words)),
            }
        )
    return output



def extract_catalog_rows(pdf_path: str | Path) -> Tuple[List[CatalogRow], List[str]]:
    rows: List[CatalogRow] = []
    warnings: List[str] = []
    category = ""
    supplier_brand = "NUVO-PRO"

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            pending_pre_lines: List[str] = []
            current: Optional[Dict[str, object]] = None

            for line in _group_page_lines(page):
                text = line["text"]
                upper = text.upper()
                if not text:
                    continue

                if "BRAKE SHOE" in upper:
                    category = "Brake Shoe"
                    pending_pre_lines = []
                    current = None
                    continue
                if "BRAKE PAD" in upper:
                    category = "Brake Pad"
                    pending_pre_lines = []
                    current = None
                    continue
                if "NUVO-PRO" in upper:
                    supplier_brand = "NUVO-PRO"
                    continue
                if any(term in upper for term in NOISE_TERMS):
                    continue

                code_match = CODE_REGEX.match(upper)
                if code_match:
                    if current:
                        rows.append(
                            CatalogRow(
                                page=int(current["page"]),
                                supplier_brand=str(current["supplier_brand"]),
                                category=str(current["category"]),
                                code=str(current["code"]),
                                axle=str(current["axle"]),
                                description=_clean_text(str(current["description"])),
                                price=current.get("price"),
                                source_line=_clean_text(f"{current['code']} {current['axle']} {current['description']}") if current.get("axle") else _clean_text(f"{current['code']} {current['description']}") ,
                            )
                        )
                        current = None

                    code = code_match.group("code").upper()
                    rest = _clean_text(code_match.group("rest"))
                    axle = ""
                    axle_match = AXLE_REGEX.match(rest.upper())
                    if axle_match:
                        axle = _normalize_axle(axle_match.group("axle"))
                        rest = _clean_text(axle_match.group("rest"))

                    description = " ".join(pending_pre_lines + ([rest] if rest else []))
                    pending_pre_lines = []
                    current = {
                        "page": page_number,
                        "supplier_brand": supplier_brand,
                        "category": category or "Uncategorized",
                        "code": code,
                        "axle": axle,
                        "description": description,
                        "price": None,
                    }

                    price_match = PRICE_REGEX.search(description)
                    if price_match:
                        current["price"] = float(price_match.group("price").replace(",", ""))
                        current["description"] = _clean_text(description[: price_match.start()])
                        rows.append(
                            CatalogRow(
                                page=page_number,
                                supplier_brand=supplier_brand,
                                category=str(current["category"]),
                                code=code,
                                axle=axle,
                                description=_clean_text(str(current["description"])),
                                price=current.get("price"),
                                source_line=_clean_text(f"{code} {axle} {current['description']}") if axle else _clean_text(f"{code} {current['description']}"),
                            )
                        )
                        current = None
                    continue

                if current is not None:
                    current["description"] = _clean_text(f"{current['description']} {text}")
                    price_match = PRICE_REGEX.search(str(current["description"]))
                    if price_match:
                        current["price"] = float(price_match.group("price").replace(",", ""))
                        current["description"] = _clean_text(str(current["description"])[: price_match.start()])
                        rows.append(
                            CatalogRow(
                                page=int(current["page"]),
                                supplier_brand=str(current["supplier_brand"]),
                                category=str(current["category"]),
                                code=str(current["code"]),
                                axle=str(current["axle"]),
                                description=_clean_text(str(current["description"])),
                                price=current.get("price"),
                                source_line=_clean_text(f"{current['code']} {current['axle']} {current['description']}") if current.get("axle") else _clean_text(f"{current['code']} {current['description']}") ,
                            )
                        )
                        current = None
                    continue

                if float(line["min_x"]) > 110:
                    pending_pre_lines.append(text)

            if current is not None:
                rows.append(
                    CatalogRow(
                        page=int(current["page"]),
                        supplier_brand=str(current["supplier_brand"]),
                        category=str(current["category"]),
                        code=str(current["code"]),
                        axle=str(current["axle"]),
                        description=_clean_text(str(current["description"])),
                        price=current.get("price"),
                        source_line=_clean_text(f"{current['code']} {current['axle']} {current['description']}") if current.get("axle") else _clean_text(f"{current['code']} {current['description']}"),
                    )
                )

    if not rows:
        warnings.append("No catalog rows were detected. This parser works best with text-based automotive catalog PDFs.")
    return rows, warnings


# -----------------------------
# Row expansion / smart parsing
# -----------------------------


def _normalize_axle(value: str) -> str:
    text = _clean_text(value).upper().replace(" ", "")
    if text in {"FRONT", "FR", "FRT"}:
        return "Front"
    if text in {"REAR", "RR"}:
        return "Rear"
    if "RR" in text and ("FRT" in text or "FRONT" in text or text.startswith("FR")):
        return "Front / Rear"
    return _clean_text(value).title()



def _split_segments(description: str) -> List[str]:
    text = description.replace(" / ", "/").replace("/ ", "/").replace(" /", "/")
    raw_segments = [segment.strip(" -") for segment in text.split("/") if segment.strip(" -")]
    merged: List[str] = []
    for segment in raw_segments:
        if merged and re.match(
            r"^(?:\d+(?:V|MM|\"|X\d+)?|'?\d{2}(?:\s*-\s*'?\d{2}|-UP)?|\(?\d.*\)?|ORIG TYPE\b|W/|DOUBLE TIRE\b)",
            segment,
            re.IGNORECASE,
        ):
            merged[-1] = f"{merged[-1]} / {segment}"
        else:
            merged.append(segment)
    return merged



def _parse_segment(segment: str, carry_brand: str = "") -> Dict[str, object]:
    raw = _cleanup_model(segment)
    brand = carry_brand
    confidence = 0.55
    notes: List[str] = []

    brand_match = BRAND_REGEX.match(raw)
    if brand_match:
        brand = BRAND_ALIASES[brand_match.group("brand").upper()]
        raw = _cleanup_model(brand_match.group("rest"))
        confidence += 0.20
        notes.append("brand token")
    elif carry_brand:
        confidence += 0.20
        notes.append("brand carried forward")

    year, raw = _extract_year(raw)
    if year:
        year = _normalize_year(year)
        confidence += 0.15
        notes.append("year parsed")

    engine_tokens = [match.group(0).upper().replace("C-240", "C240") for match in ENGINE_REGEX.finditer(raw)]
    unique_engine_tokens: List[str] = []
    for token in engine_tokens:
        if token not in unique_engine_tokens:
            unique_engine_tokens.append(token)
    engine = ", ".join(unique_engine_tokens)
    if unique_engine_tokens:
        confidence += 0.08
        notes.append("engine code parsed")
        for token in unique_engine_tokens:
            raw = re.sub(rf"\b{re.escape(token.replace('C240', 'C-240'))}\b", " ", raw, flags=re.IGNORECASE)
            raw = re.sub(rf"\b{re.escape(token)}\b", " ", raw, flags=re.IGNORECASE)
        raw = _cleanup_model(raw)

    return {
        "brand": brand,
        "model": raw,
        "engine": engine,
        "year": year,
        "confidence": round(min(confidence, 0.99), 2),
        "pattern_notes": ", ".join(notes),
    }



def _build_model_brand_map(application_rows: Iterable[Dict[str, object]]) -> Dict[str, str]:
    counter: Counter[Tuple[str, str]] = Counter()
    for row in application_rows:
        brand = _clean_text(row.get("Brand", "")).upper()
        model = _clean_text(row.get("Model", "")).upper()
        if not brand or not model:
            continue
        first_word = _model_key(model).split(" ")[0] if _model_key(model) else ""
        if first_word:
            counter[(first_word, brand)] += 1

    mapping: Dict[str, str] = {}
    grouped: Dict[str, List[Tuple[str, int]]] = defaultdict(list)
    for (first_word, brand), count in counter.items():
        grouped[first_word].append((brand, count))

    for first_word, values in grouped.items():
        values.sort(key=lambda x: (-x[1], x[0]))
        mapping[first_word] = values[0][0]
    return mapping



def expand_catalog_rows(raw_rows: List[CatalogRow]) -> Tuple[List[Dict[str, object]], List[Dict[str, object]]]:
    expanded: List[Dict[str, object]] = []

    for row in raw_rows:
        carry_brand = ""
        segments = _split_segments(row.description)
        for segment_index, segment in enumerate(segments, start=1):
            parsed = _parse_segment(segment, carry_brand=carry_brand)
            if parsed["brand"]:
                carry_brand = str(parsed["brand"])

            expanded.append(
                {
                    "Source Line": row.source_line,
                    "Supplier Brand": row.supplier_brand,
                    "Category": row.category,
                    "Code": row.code,
                    "Axle": _normalize_axle(row.axle),
                    "Brand": parsed["brand"],
                    "Model": parsed["model"],
                    "Engine": parsed["engine"],
                    "Year": parsed["year"],
                    "Original Price (PHP)": row.price,
                    "Your Price (PHP)": None,
                    "Use Price (PHP)": None,
                    "Page": row.page,
                    "Confidence": parsed["confidence"],
                    "Review Status": "",
                    "Pattern Notes": parsed["pattern_notes"],
                    "_Segment": segment,
                    "_Segment Order": segment_index,
                }
            )

    model_brand_map = _build_model_brand_map(expanded)

    review_rows: List[Dict[str, object]] = []
    for row in expanded:
        if not row["Brand"]:
            model_key = _model_key(str(row["Model"]))
            first_word = model_key.split(" ")[0] if model_key else ""
            inferred_brand = model_brand_map.get(first_word, "")
            if inferred_brand:
                row["Brand"] = inferred_brand
                row["Confidence"] = round(min(float(row["Confidence"]) + 0.12, 0.95), 2)
                row["Pattern Notes"] = (
                    f"{row['Pattern Notes']}, brand inferred from model map ({inferred_brand})"
                    if row["Pattern Notes"]
                    else f"brand inferred from model map ({inferred_brand})"
                )

        # Obvious missing details / confidence labels
        confidence = float(row["Confidence"])
        if not row["Brand"] or not row["Model"]:
            status = "Needs Review"
            confidence = min(confidence, 0.60)
        elif confidence >= 0.90:
            status = "Auto-Accepted"
        elif confidence >= 0.75:
            status = "Review Suggested"
        else:
            status = "Needs Review"

        row["Confidence"] = round(confidence, 2)
        row["Review Status"] = status

        if status != "Auto-Accepted":
            review_rows.append({key: row[key] for key in MAIN_COLUMNS})

    expanded.sort(key=lambda x: (int(x["Page"]), str(x["Code"]), int(x["_Segment Order"])))
    return expanded, review_rows



def parse_catalog_pdf(pdf_path: str | Path) -> ParsedCatalog:
    raw_rows, warnings = extract_catalog_rows(pdf_path)
    application_rows, review_rows = expand_catalog_rows(raw_rows)
    return ParsedCatalog(raw_rows=raw_rows, application_rows=application_rows, review_rows=review_rows, warnings=warnings)


# -----------------------------
# DataFrame helpers
# -----------------------------


def applications_dataframe(rows: List[Dict[str, object]]) -> pd.DataFrame:
    return pd.DataFrame([{column: row.get(column, "") for column in MAIN_COLUMNS} for row in rows], columns=MAIN_COLUMNS)



def raw_rows_dataframe(rows: List[CatalogRow]) -> pd.DataFrame:
    data = []
    for row in rows:
        data.append(
            {
                "Page": row.page,
                "Supplier Brand": row.supplier_brand,
                "Category": row.category,
                "Code": row.code,
                "Axle": row.axle,
                "Description": row.description,
                "Original Price (PHP)": row.price,
            }
        )
    return pd.DataFrame(data, columns=RAW_COLUMNS)


# -----------------------------
# Excel export
# -----------------------------


def _write_df_to_sheet(wb: Workbook, title: str, df: pd.DataFrame) -> None:
    ws = wb.create_sheet(title=title)
    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for col_idx, column in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=column)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    your_price_col = df.columns.get_loc("Your Price (PHP)") + 1 if "Your Price (PHP)" in df.columns else None
    original_price_col = df.columns.get_loc("Original Price (PHP)") + 1 if "Original Price (PHP)" in df.columns else None
    use_price_col = df.columns.get_loc("Use Price (PHP)") + 1 if "Use Price (PHP)" in df.columns else None

    for row_idx, (_, row) in enumerate(df.iterrows(), start=2):
        for col_idx, column in enumerate(df.columns, start=1):
            value = row[column]
            if pd.isna(value):
                value = None
            if column == "Use Price (PHP)" and your_price_col and original_price_col:
                your_ref = f"{get_column_letter(your_price_col)}{row_idx}"
                original_ref = f"{get_column_letter(original_price_col)}{row_idx}"
                ws.cell(row=row_idx, column=col_idx, value=f'=IF({your_ref}="",{original_ref},{your_ref})')
            else:
                ws.cell(row=row_idx, column=col_idx, value=value)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for col_idx, column in enumerate(df.columns, start=1):
        if column in {"Original Price (PHP)", "Your Price (PHP)", "Use Price (PHP)"}:
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row=row_idx, column=col_idx).number_format = '₱#,##0.00'
        if column == "Confidence":
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row=row_idx, column=col_idx).number_format = '0%'

    for col_idx, column in enumerate(df.columns, start=1):
        max_len = max(len(_clean_text(column)), *(len(_clean_text(ws.cell(r, col_idx).value)) for r in range(2, min(ws.max_row, 250) + 1)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 10), 42)



def build_workbook(applications_df: pd.DataFrame, raw_df: pd.DataFrame, review_df: Optional[pd.DataFrame] = None) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    _write_df_to_sheet(wb, "Applications", applications_df)
    _write_df_to_sheet(wb, "Catalog Rows", raw_df)
    _write_df_to_sheet(wb, "Review Queue", review_df if review_df is not None else applications_df.iloc[0:0])
    return wb



def workbook_to_bytes(wb: Workbook) -> bytes:
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()



def build_demo_excel(pdf_path: str | Path, output_path: str | Path) -> None:
    parsed = parse_catalog_pdf(pdf_path)
    applications_df = applications_dataframe(parsed.application_rows)
    raw_df = raw_rows_dataframe(parsed.raw_rows)
    review_df = pd.DataFrame(parsed.review_rows, columns=MAIN_COLUMNS)
    wb = build_workbook(applications_df, raw_df, review_df)
    wb.save(str(output_path))
