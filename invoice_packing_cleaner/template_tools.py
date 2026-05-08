from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import pandas as pd

from invoice_packing_cleaner.table_tools import clean_header, excel_column_label, make_unique_headers


HEADER_KEYWORDS = {
    "no",
    "inv",
    "invoice",
    "marks",
    "nos",
    "po",
    "part",
    "item",
    "description",
    "desc",
    "goods",
    "qty",
    "quantity",
    "unit",
    "price",
    "amount",
    "nw",
    "gw",
    "weight",
    "ctn",
    "carton",
    "carton no",
    "ctn no",
    "package",
    "measurement",
    "cbm",
    "hs",
    "brand",
    "chk",
    "品名",
    "數量",
    "單位",
    "單價",
    "金額",
    "毛重",
    "淨重",
    "箱",
    "件",
    "稅則",
}


@dataclass(frozen=True)
class TemplateColumn:
    name: str
    column_index: int

    @property
    def column_letter(self) -> str:
        return excel_column_label(self.column_index)


@dataclass(frozen=True)
class TemplateCandidate:
    label: str
    file_name: str
    sheet_name: str
    header_row: int
    data_start_row: int
    columns: list[TemplateColumn]
    score: int
    dataframe: pd.DataFrame


def parse_output_template_file(file_name: str, data: bytes) -> list[TemplateCandidate]:
    suffix = Path(file_name).suffix.lower()

    if suffix in {".xlsx", ".xls", ".xlsm"}:
        return _parse_excel_template(file_name, data)
    if suffix == ".csv":
        return _parse_csv_template(file_name, data)

    raise ValueError("最終格式範本目前請上傳 Excel 或 CSV。")


def template_columns_to_text(columns: list[TemplateColumn]) -> str:
    return "\n".join(column.name for column in columns)


def build_template_preview(columns: list[TemplateColumn], header_row: int) -> pd.DataFrame:
    rows = []
    for column in columns:
        rows.append(
            {
                "輸出欄位": column.name,
                "偵測位置": f"{column.column_letter}{header_row}",
                "輸出欄號": column.column_index,
            }
        )
    return pd.DataFrame(rows)


def _parse_excel_template(file_name: str, data: bytes) -> list[TemplateCandidate]:
    buffer = BytesIO(data)
    workbook = pd.ExcelFile(buffer)
    candidates: list[TemplateCandidate] = []

    for sheet_name in workbook.sheet_names:
        buffer.seek(0)
        df = pd.read_excel(
            buffer,
            sheet_name=sheet_name,
            header=None,
            dtype=str,
            keep_default_na=False,
        )
        sheet_candidates = _detect_candidates(file_name, sheet_name, df)
        if not sheet_candidates:
            fallback = _fallback_candidate(file_name, sheet_name, df)
            if fallback:
                sheet_candidates.append(fallback)
        candidates.extend(sheet_candidates)

    return candidates


def _parse_csv_template(file_name: str, data: bytes) -> list[TemplateCandidate]:
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
            return _detect_candidates(file_name, "CSV", df)
        except Exception as exc:  # pragma: no cover - depends on file encoding
            last_error = exc

    raise ValueError(f"最終格式 CSV 讀取失敗：{last_error}")


def _detect_candidates(file_name: str, sheet_name: str, df: pd.DataFrame) -> list[TemplateCandidate]:
    candidates: list[TemplateCandidate] = []
    if df.empty:
        return candidates

    max_scan_rows = min(len(df), 80)
    for row_idx in range(max_scan_rows):
        raw_values = df.iloc[row_idx].tolist()
        values = [clean_header(value) for value in raw_values]
        nonblank = [(idx + 1, value) for idx, value in enumerate(values) if value]
        if len(nonblank) < 2:
            continue

        score = _score_header_row(nonblank)
        if score < 4:
            continue

        unique_headers = make_unique_headers([value for _, value in nonblank])
        columns = [
            TemplateColumn(name=header, column_index=col_idx)
            for (col_idx, _), header in zip(nonblank, unique_headers)
        ]
        label = (
            f"{file_name} / {sheet_name} / 第 {row_idx + 1} 列 "
            f"({len(columns)} 欄，分數 {score})"
        )
        candidates.append(
            TemplateCandidate(
                label=label,
                file_name=file_name,
                sheet_name=sheet_name,
                header_row=row_idx + 1,
                data_start_row=row_idx + 2,
                columns=columns,
                score=score,
                dataframe=df,
            )
        )

    candidates.sort(key=lambda candidate: candidate.score, reverse=True)
    return candidates[:10]


def _fallback_candidate(file_name: str, sheet_name: str, df: pd.DataFrame) -> TemplateCandidate | None:
    if df.empty:
        return None

    best_row_idx = -1
    best_score = -999
    best_nonblank: list[tuple[int, str]] = []
    max_scan_rows = min(len(df), 80)

    for row_idx in range(max_scan_rows):
        values = [clean_header(value) for value in df.iloc[row_idx].tolist()]
        nonblank = [(idx + 1, value) for idx, value in enumerate(values) if value]
        if len(nonblank) < 2:
            continue

        score = _score_header_row(nonblank)
        if score > best_score:
            best_row_idx = row_idx
            best_score = score
            best_nonblank = nonblank

    if best_row_idx < 0 or not best_nonblank:
        return None

    unique_headers = make_unique_headers([value for _, value in best_nonblank])
    columns = [
        TemplateColumn(name=header, column_index=col_idx)
        for (col_idx, _), header in zip(best_nonblank, unique_headers)
    ]

    return TemplateCandidate(
        label=(
            f"{file_name} / {sheet_name} / 第 {best_row_idx + 1} 列 "
            f"({len(columns)} 欄，備用偵測，分數 {best_score})"
        ),
        file_name=file_name,
        sheet_name=sheet_name,
        header_row=best_row_idx + 1,
        data_start_row=best_row_idx + 2,
        columns=columns,
        score=best_score,
        dataframe=df,
    )


def _score_header_row(nonblank: list[tuple[int, str]]) -> int:
    score = len(nonblank)
    numeric_like = 0
    keyword_hits = 0

    for _, value in nonblank:
        normalized = value.lower().replace(".", "").replace("#", "")
        if _is_numeric_like(value):
            numeric_like += 1
        if any(keyword in normalized for keyword in HEADER_KEYWORDS):
            keyword_hits += 1

    score += keyword_hits * 3

    if keyword_hits >= 2:
        score += 4
    if numeric_like >= max(2, len(nonblank) // 2):
        score -= 5
    if any("total" in value.lower() or "合計" in value for _, value in nonblank):
        score -= 2

    return score


def _is_numeric_like(value: str) -> bool:
    compact = value.replace(",", "").replace(".", "").replace("-", "").strip()
    return bool(compact) and compact.isdigit()
