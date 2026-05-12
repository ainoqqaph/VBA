from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd


DEFAULT_TARGET_COLUMNS = [
    "Invoice No",
    "Invoice Date",
    "Shipper",
    "Consignee",
    "Item No",
    "Description",
    "HS Code",
    "Qty",
    "Unit",
    "Unit Price",
    "Amount",
    "Currency",
    "Net Weight",
    "Gross Weight",
    "Packages",
    "Package Unit",
    "Measurement",
    "Country of Origin",
]


OP_TINV_COLUMNS = [
    "PO No.",
    "Item No",
    "Description",
    "Description 2",
    "Description 3",
    "Description 4",
    "Description 5",
    "Line No.",
    "Quantity",
    "Unit",
    "Unit Price",
    "Amount",
    "HS Code",
    "Brand",
    "Customer Item",
    "Title",
]

OP_TPKG_COLUMNS = [
    "Customer PO",
    "PO No.",
    "Carton No",
    "Carton From",
    "Carton To",
    "CTN",
    "Item No",
    "Description",
    "Description 2",
    "Description 3",
    "Description 4",
    "Description 5",
    "Quantity",
    "Unit Qty",
    "Unit",
    "Net Weight",
    "Gross Weight",
    "Measurement",
    "Title",
]


KEYWORD_MAP = {
    "Invoice No": ["invoice no", "inv no", "invoice#", "invoice number", "發票號碼"],
    "Invoice Date": ["invoice date", "date", "發票日期"],
    "Shipper": ["shipper", "seller", "exporter", "供應商", "出口商"],
    "Consignee": ["consignee", "buyer", "importer", "收貨人", "買方", "進口商"],
    "Item No": ["item", "item no", "part no", "model", "料號", "型號", "項次"],
    "Description": ["description", "goods", "commodity", "product", "品名", "貨名"],
    "HS Code": ["hs code", "hscode", "hs", "稅則", "稅號"],
    "Qty": ["qty", "quantity", "數量"],
    "Unit": ["unit", "uom", "單位"],
    "Unit Price": ["unit price", "price", "單價"],
    "Amount": ["amount", "total", "金額", "總價"],
    "Currency": ["currency", "幣別"],
    "Net Weight": ["net weight", "nw", "n.w.", "淨重"],
    "Gross Weight": ["gross weight", "gw", "g.w.", "毛重"],
    "Packages": ["packages", "package", "carton", "ctn", "箱數", "件數"],
    "Package Unit": ["package unit", "pkg unit", "包裝單位"],
    "Measurement": ["measurement", "cbm", "volume", "材積"],
    "Country of Origin": ["origin", "country of origin", "coo", "產地", "原產地"],
}


KEYWORD_MAP.update(
    {
        "PO No.": ["po", "po no", "p/o", "order", "order no", "order#"],
        "Customer PO": ["cust po", "customer po", "po", "po no", "p/o", "order", "order#"],
        "Item No": [
            "item",
            "item no",
            "item#",
            "part no",
            "partino",
            "caprparts",
            "caprparts#",
            "carpart",
            "carparts",
            "model",
        ],
        "Line No.": ["line", "line no"],
        "Description 2": ["description 2", "partino", "part no", "customer#", "customer no", "ys#", "oe#", "pls#", "model"],
        "Description 3": ["description 3", "customer#", "customer no", "ys#", "oe#", "pls#", "model"],
        "Description 4": ["description 4", "ys#", "oe#", "pls#", "model"],
        "Description 5": ["description 5", "oe#", "pls#", "model"],
        "Customer Item": ["customer item", "customer#", "customer no", "cust item", "cust part"],
        "Quantity": ["qty", "quantity", "q'ty", "數量"],
        "Carton No": ["packing no", "carton no", "ctn no", "case no", "marks", "marks nos"],
        "Carton From": ["carton from", "ctn from", "from carton", "start carton", "start"],
        "Carton To": ["carton to", "ctn to", "to carton", "end carton", "end"],
        "CTN": ["ctn", "ct.no", "ct no", "carton", "cartons", "packages"],
        "Unit Qty": ["unit qty", "pcs/carton", "pc/ctn", "pcs/ctn", "q'ty ctn", "qty/ctn", "pac."],
        "Title": ["title", "category", "invoice of"],
    }
)


@dataclass(frozen=True)
class SourceColumn:
    label: str
    index: int
    header: str


def excel_column_label(index: int) -> str:
    if index < 1:
        return ""
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def split_target_columns(text: str) -> list[str]:
    columns: list[str] = []
    for raw_line in text.replace(",", "\n").splitlines():
        column = raw_line.strip()
        if column:
            columns.append(column)
    return columns


def prepare_structured_table(
    raw_df: pd.DataFrame,
    header_row: int,
    data_start_row: int,
) -> tuple[pd.DataFrame, list[str]]:
    if raw_df.empty:
        return raw_df.copy(), []

    header_index = max(header_row - 1, 0)
    data_index = max(data_start_row - 1, header_index + 1)
    raw_headers = raw_df.iloc[header_index].tolist()
    headers = make_unique_headers(raw_headers)
    headers = _rename_blank_carton_range_headers(raw_headers, headers)
    structured = raw_df.iloc[data_index:].copy()
    structured.columns = headers[: structured.shape[1]]
    structured = structured.loc[~structured.apply(_is_blank_row, axis=1)]
    return structured, headers


def _rename_blank_carton_range_headers(raw_headers: list[object], headers: list[str]) -> list[str]:
    renamed = list(headers)
    last_named = ""
    blank_after_total = 0

    for index, raw_header in enumerate(raw_headers):
        header = clean_header(raw_header)
        if header:
            last_named = header
            blank_after_total = 0
            continue

        if _normalize_lookup(last_named) != "total":
            continue

        blank_after_total += 1
        if blank_after_total == 1:
            renamed[index] = "Carton From"
        elif blank_after_total == 2:
            renamed[index] = "Carton To"

    return renamed


def suggest_header_row(raw_df: pd.DataFrame) -> int:
    if raw_df.empty:
        return 1

    best_row = 1
    best_score = -10**9
    max_scan_rows = min(len(raw_df), 90)
    keyword_values = [keyword for values in KEYWORD_MAP.values() for keyword in values]

    for row_idx in range(max_scan_rows):
        values = [clean_header(value) for value in raw_df.iloc[row_idx].tolist()]
        nonblank = [value for value in values if value]
        if len(nonblank) < 2:
            continue

        normalized_values = [_normalize_lookup(value) for value in nonblank]
        keyword_hits = 0
        for value in normalized_values:
            if any(_normalize_lookup(keyword) and _normalize_lookup(keyword) in value for keyword in keyword_values):
                keyword_hits += 1

        numeric_like = sum(_is_numeric_like(value) for value in nonblank)
        score = len(nonblank) + keyword_hits * 4

        if keyword_hits >= 3:
            score += 8
        if any("seq" in value or "no" == value for value in normalized_values):
            score += 2
        if numeric_like >= max(2, len(nonblank) // 2):
            score -= 10
        if any("total" in value for value in normalized_values):
            score -= 4

        if score > best_score:
            best_score = score
            best_row = row_idx + 1

    return best_row


def make_unique_headers(raw_headers: list[object]) -> list[str]:
    headers: list[str] = []
    counts: dict[str, int] = {}
    last_named_header = ""

    for index, raw_header in enumerate(raw_headers, start=1):
        header = clean_header(raw_header)
        if not header:
            if last_named_header:
                header = f"{last_named_header} {excel_column_label(index)}"
            else:
                header = f"Column {excel_column_label(index)}"
        else:
            last_named_header = header

        base = header
        counts[base] = counts.get(base, 0) + 1
        if counts[base] > 1:
            header = f"{base}_{counts[base]}"

        headers.append(header)

    return headers


def build_source_options(headers: list[str]) -> list[SourceColumn]:
    options = [SourceColumn("（留空）", 0, "")]
    for index, header in enumerate(headers, start=1):
        options.append(SourceColumn(_source_option_label(index, header), index, header))
    return options


def _source_option_label(index: int, header: str) -> str:
    preview_index = str(index - 1)
    header = str(header or "").strip()

    if not header or _is_generated_column_header(index, header):
        return preview_index

    short_header = header if len(header) <= 70 else f"{header[:67]}..."
    return f"{preview_index} | {short_header}"


def _is_generated_column_header(index: int, header: str) -> bool:
    column_letter = excel_column_label(index)
    return bool(re.fullmatch(rf"Column {re.escape(column_letter)}(?:_\d+)?", header))


def find_default_source_index(target: str, options: list[SourceColumn]) -> int:
    keywords = KEYWORD_MAP.get(target, [target])
    normalized_headers = [
        (position, _normalize_lookup(option.header))
        for position, option in enumerate(options)
        if option.index > 0
    ]
    description_positions = [
        position
        for position, normalized_header in normalized_headers
        if "description" in normalized_header or "desc" in normalized_header
    ]
    description_family_positions = [
        position
        for position, normalized_header in normalized_headers
        if any(
            token in normalized_header
            for token in ("description", "desc", "partino", "partno", "customer", "model", "ys", "oe", "pls")
        )
    ]

    if target == "Item No" and description_positions:
        item_like = [
            position
            for position, normalized_header in normalized_headers
            if any(token in normalized_header for token in ("item", "part", "caprpart", "carpart", "model"))
        ]
        if not item_like:
            return description_positions[0]

    if target == "Description" and description_family_positions:
        return description_family_positions[0]

    if target.startswith("Description ") and (description_positions or description_family_positions):
        try:
            offset = int(target.split()[-1]) - 1
        except ValueError:
            offset = 0
        if 0 <= offset < len(description_family_positions):
            return description_family_positions[offset]
        if 0 <= offset < len(description_positions):
            return description_positions[offset]

    if target == "Carton No":
        has_carton_range = any(
            normalized_header in {"cartonfrom", "cartonto", "ctnfrom", "ctnto", "start", "end"}
            for _, normalized_header in normalized_headers
        )
        if has_carton_range:
            return 0

    for keyword in keywords:
        normalized_keyword = _normalize_lookup(keyword)
        if not normalized_keyword:
            continue
        for position, normalized_header in normalized_headers:
            if normalized_keyword == normalized_header:
                return position
        for position, normalized_header in normalized_headers:
            if normalized_keyword in normalized_header or (
                len(normalized_header) >= 4 and normalized_header in normalized_keyword
            ):
                return position

    return 0


def clean_header(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return " ".join(text.split()).strip()


def _normalize_lookup(value: object) -> str:
    text = clean_header(value).lower()
    return re.sub(r"[\s_\-./:#()（）'&]+", "", text)


def _is_blank_row(row: pd.Series) -> bool:
    return all(clean_header(value) == "" for value in row.tolist())


def _is_numeric_like(value: str) -> bool:
    compact = clean_header(value)
    compact = compact.replace(",", "").replace(".", "").replace("-", "")
    return bool(compact) and compact.isdigit()
