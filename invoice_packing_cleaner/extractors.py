from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import pandas as pd


@dataclass(frozen=True)
class ParsedTable:
    label: str
    file_name: str
    kind: str
    dataframe: pd.DataFrame
    note: str = ""


class UnsupportedFileType(ValueError):
    pass


class MissingDependency(RuntimeError):
    pass


def parse_uploaded_file(file_name: str, data: bytes) -> list[ParsedTable]:
    suffix = Path(file_name).suffix.lower()

    if suffix in {".xlsx", ".xls", ".xlsm"}:
        return _parse_excel(file_name, data)
    if suffix == ".csv":
        return _parse_csv(file_name, data)
    if suffix == ".pdf":
        return _parse_pdf(file_name, data)
    if suffix == ".docx":
        return _parse_docx(file_name, data)
    if suffix == ".doc":
        raise UnsupportedFileType(
            "舊版 .doc 需要先另存成 .docx，或用 Word/LibreOffice 轉檔後再上傳。"
        )

    raise UnsupportedFileType(f"目前不支援 {suffix or '未知'} 檔案格式。")


def _normalize_cell(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value)
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return " ".join(text.split())


def _normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    normalized = df.fillna("").astype(str)
    return normalized.apply(lambda col: col.map(_normalize_cell))


def _parse_excel(file_name: str, data: bytes) -> list[ParsedTable]:
    buffer = BytesIO(data)
    workbook = pd.ExcelFile(buffer)
    tables: list[ParsedTable] = []

    for sheet_name in workbook.sheet_names:
        buffer.seek(0)
        df = pd.read_excel(
            buffer,
            sheet_name=sheet_name,
            header=None,
            dtype=str,
            keep_default_na=False,
        )
        tables.append(
            ParsedTable(
                label=f"{file_name} / {sheet_name}",
                file_name=file_name,
                kind="excel",
                dataframe=_normalize_dataframe(df),
            )
        )

    return tables


def _parse_csv(file_name: str, data: bytes) -> list[ParsedTable]:
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "cp950", "big5", "latin1"):
        try:
            df = pd.read_csv(
                BytesIO(data),
                header=None,
                dtype=str,
                keep_default_na=False,
                encoding=encoding,
            )
            return [
                ParsedTable(
                    label=f"{file_name} / CSV",
                    file_name=file_name,
                    kind="csv",
                    dataframe=_normalize_dataframe(df),
                    note=f"使用 {encoding} 編碼讀取。",
                )
            ]
        except Exception as exc:  # pragma: no cover - depends on input encoding
            last_error = exc

    raise ValueError(f"CSV 讀取失敗：{last_error}")


def _parse_pdf(file_name: str, data: bytes) -> list[ParsedTable]:
    try:
        import pdfplumber
    except ModuleNotFoundError as exc:  # pragma: no cover - dependency guard
        raise MissingDependency("請先安裝 pdfplumber：pip install pdfplumber") from exc

    tables: list[ParsedTable] = []

    with pdfplumber.open(BytesIO(data)) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            extracted_tables = page.extract_tables() or []
            if extracted_tables:
                for table_number, table in enumerate(extracted_tables, start=1):
                    rows = _pad_rows(table)
                    if rows:
                        tables.append(
                            ParsedTable(
                                label=f"{file_name} / PDF 第 {page_number} 頁 表格 {table_number}",
                                file_name=file_name,
                                kind="pdf-table",
                                dataframe=_normalize_dataframe(pd.DataFrame(rows)),
                            )
                        )
                continue

            text = page.extract_text() or ""
            lines = [[line] for line in text.splitlines() if line.strip()]
            if lines:
                tables.append(
                    ParsedTable(
                        label=f"{file_name} / PDF 第 {page_number} 頁 文字",
                        file_name=file_name,
                        kind="pdf-text",
                        dataframe=_normalize_dataframe(pd.DataFrame(lines)),
                        note="這頁沒有偵測到表格，已用逐行文字方式讀取。",
                    )
                )

    if not tables:
        tables.append(
            ParsedTable(
                label=f"{file_name} / PDF",
                file_name=file_name,
                kind="pdf-empty",
                dataframe=pd.DataFrame(),
                note="沒有讀到可用文字或表格；若是掃描圖檔 PDF，之後需要加 OCR。",
            )
        )

    return tables


def _parse_docx(file_name: str, data: bytes) -> list[ParsedTable]:
    try:
        from docx import Document
    except ModuleNotFoundError as exc:  # pragma: no cover - dependency guard
        raise MissingDependency("請先安裝 python-docx：pip install python-docx") from exc

    document = Document(BytesIO(data))
    tables: list[ParsedTable] = []

    for table_number, table in enumerate(document.tables, start=1):
        rows = [[cell.text for cell in row.cells] for row in table.rows]
        rows = _pad_rows(rows)
        if rows:
            tables.append(
                ParsedTable(
                    label=f"{file_name} / Word 表格 {table_number}",
                    file_name=file_name,
                    kind="word-table",
                    dataframe=_normalize_dataframe(pd.DataFrame(rows)),
                )
            )

    if tables:
        return tables

    lines = [[paragraph.text] for paragraph in document.paragraphs if paragraph.text.strip()]
    if lines:
        return [
            ParsedTable(
                label=f"{file_name} / Word 文字",
                file_name=file_name,
                kind="word-text",
                dataframe=_normalize_dataframe(pd.DataFrame(lines)),
                note="這份 Word 沒有表格，已用逐段文字方式讀取。",
            )
        ]

    return [
        ParsedTable(
            label=f"{file_name} / Word",
            file_name=file_name,
            kind="word-empty",
            dataframe=pd.DataFrame(),
            note="沒有讀到可用文字或表格。",
        )
    ]


def _pad_rows(rows: list[list[object] | None]) -> list[list[object]]:
    clean_rows = [list(row or []) for row in rows]
    if not clean_rows:
        return []
    max_len = max(len(row) for row in clean_rows)
    return [row + [""] * (max_len - len(row)) for row in clean_rows]
