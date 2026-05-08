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
    structured = raw_df.iloc[data_index:].copy()
    structured.columns = headers[: structured.shape[1]]
    structured = structured.loc[~structured.apply(_is_blank_row, axis=1)]
    return structured, headers


def make_unique_headers(raw_headers: list[object]) -> list[str]:
    headers: list[str] = []
    counts: dict[str, int] = {}

    for index, raw_header in enumerate(raw_headers, start=1):
        header = clean_header(raw_header)
        if not header:
            header = f"Column {excel_column_label(index)}"

        base = header
        counts[base] = counts.get(base, 0) + 1
        if counts[base] > 1:
            header = f"{base}_{counts[base]}"

        headers.append(header)

    return headers


def build_source_options(headers: list[str]) -> list[SourceColumn]:
    options = [SourceColumn("（留空）", 0, "")]
    for index, header in enumerate(headers, start=1):
        column_letter = excel_column_label(index)
        short_header = header if len(header) <= 70 else f"{header[:67]}..."
        options.append(SourceColumn(f"{column_letter} | {short_header}", index, header))
    return options


def find_default_source_index(target: str, options: list[SourceColumn]) -> int:
    keywords = KEYWORD_MAP.get(target, [target])
    normalized_headers = [
        (position, _normalize_lookup(option.header))
        for position, option in enumerate(options)
        if option.index > 0
    ]

    for keyword in keywords:
        normalized_keyword = _normalize_lookup(keyword)
        if not normalized_keyword:
            continue
        for position, normalized_header in normalized_headers:
            if normalized_keyword == normalized_header:
                return position
        for position, normalized_header in normalized_headers:
            if normalized_keyword in normalized_header or normalized_header in normalized_keyword:
                return position

    return 0


def clean_header(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return " ".join(text.split()).strip()


def _normalize_lookup(value: object) -> str:
    text = clean_header(value).lower()
    return re.sub(r"[\s_\-./:#()（）]+", "", text)


def _is_blank_row(row: pd.Series) -> bool:
    return all(clean_header(value) == "" for value in row.tolist())
