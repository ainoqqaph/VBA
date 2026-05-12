from __future__ import annotations

import re
from typing import Any, Iterable


def pdf_page_to_rows(page: Any) -> list[list[str]]:
    """Return text rows from a pdfplumber page while preserving table columns."""

    try:
        words = page.extract_words(
            x_tolerance=1,
            y_tolerance=3,
            keep_blank_chars=False,
            use_text_flow=False,
        )
    except TypeError:
        words = page.extract_words(x_tolerance=1, y_tolerance=3) or []
    except Exception:
        words = []

    rows = words_to_rows(words)
    if rows:
        return rows

    text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
    return text_to_rows(text)


def words_to_rows(words: Iterable[dict[str, Any]]) -> list[list[str]]:
    lines = _group_words_by_line(words)
    if not lines:
        return []

    rows: list[list[str]] = []
    active_columns: list[tuple[str, float, float]] | None = None

    for line_words in lines:
        header_columns = _header_columns_from_words(line_words)
        if header_columns:
            active_columns = header_columns
            rows.append([name for name, _, _ in header_columns])
            continue

        if active_columns:
            rows.append(_assign_words_to_columns(line_words, active_columns))
        else:
            rows.append([" ".join(str(word.get("text", "")).strip() for word in line_words).strip()])

    return _pad_rows([row for row in rows if any(cell.strip() for cell in row)])


def text_to_rows(text: str) -> list[list[str]]:
    rows: list[list[str]] = []
    for raw_line in text.splitlines():
        line = _normalize_text_cell(raw_line)
        if not line:
            continue
        known_parts = _split_known_pdf_header_line(line)
        if known_parts:
            rows.append(known_parts)
            continue
        if "|" in line:
            parts = [part.strip() for part in line.strip("|").split("|") if part.strip()]
        else:
            parts = [part.strip() for part in re.split(r"\s{2,}|\t+", line) if part.strip()]
        rows.append(parts or [line])
    return _pad_rows(rows)


def _group_words_by_line(words: Iterable[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    sorted_words = sorted(
        (word for word in words if str(word.get("text", "")).strip()),
        key=lambda word: (float(word.get("top", 0)), float(word.get("x0", 0))),
    )
    lines: list[list[dict[str, Any]]] = []
    line_tops: list[float] = []
    tolerance = 3.5

    for word in sorted_words:
        top = float(word.get("top", 0))
        target_index = -1
        for index, line_top in enumerate(line_tops):
            if abs(top - line_top) <= tolerance:
                target_index = index
                break
        if target_index < 0:
            lines.append([word])
            line_tops.append(top)
        else:
            lines[target_index].append(word)
            line_tops[target_index] = (line_tops[target_index] + top) / 2

    for line in lines:
        line.sort(key=lambda word: float(word.get("x0", 0)))
    return lines


def _header_columns_from_words(line_words: list[dict[str, Any]]) -> list[tuple[str, float, float]]:
    tokens = [_normalize_token(word.get("text", "")) for word in line_words]
    joined = " ".join(tokens)
    joined_compact = "".join(tokens)

    if "packing" in joined and "description" in joined and "quantity" in joined:
        return _find_header_sequence(
            line_words,
            (
                ("Packing No.", (("packing", "no"),)),
                ("Description Of Goods", (("description", "of", "goods"),)),
                ("Quantity", (("quantity",),)),
                ("Net Weight", (("net", "weight"), ("netweight",), ("netwe", "ight"))),
                ("Gross Weight", (("gross", "weight"), ("grossweight",))),
                ("Measurement", (("measurement",), ("me", "asurement"))),
            ),
        )

    if "description" in joined and "quantity" in joined and (
        "unitprice" in joined_compact or "amount" in joined
    ):
        return _find_header_sequence(
            line_words,
            (
                ("Marks & Nos.", (("marks", "nos"), ("marks", "no"))),
                ("Description Of Goods", (("description", "of", "goods"),)),
                ("Quantity", (("quantity",),)),
                ("Unit Price", (("unit", "price"), ("unitprice",))),
                ("Amount", (("amount",),)),
            ),
        )

    return []


def _find_header_sequence(
    line_words: list[dict[str, Any]],
    definitions: Iterable[tuple[str, Iterable[tuple[str, ...]]]],
) -> list[tuple[str, float, float]]:
    tokens = [_normalize_token(word.get("text", "")) for word in line_words]
    found: list[tuple[str, float, float]] = []
    last_index = -1

    for name, patterns in definitions:
        match_index = -1
        for pattern in patterns:
            match_index = _find_pattern(tokens, pattern, start=last_index + 1)
            if match_index >= 0:
                break
        if match_index >= 0:
            end_index = match_index + len(pattern) - 1
            found.append(
                (
                    name,
                    float(line_words[match_index].get("x0", 0)),
                    float(line_words[end_index].get("x1", line_words[end_index].get("x0", 0))),
                )
            )
            last_index = match_index

    return found if len(found) >= 3 else []


def _find_pattern(tokens: list[str], pattern: tuple[str, ...], *, start: int) -> int:
    for index in range(start, len(tokens) - len(pattern) + 1):
        candidate = tuple(token for token in tokens[index : index + len(pattern)] if token)
        if candidate == pattern:
            return index
    return -1


def _assign_words_to_columns(
    line_words: list[dict[str, Any]],
    columns: list[tuple[str, float, float]],
) -> list[str]:
    boundaries: list[float] = []
    for left, right in zip(columns, columns[1:]):
        left_name, left_x0, left_x1 = left
        _, right_x0, _ = right
        if left_name == "Description Of Goods" and left_x1 < right_x0:
            boundaries.append((left_x1 + right_x0) / 2)
        else:
            boundaries.append((left_x0 + right_x0) / 2)
    cells: list[list[str]] = [[] for _ in columns]

    for word in line_words:
        x0 = float(word.get("x0", 0))
        col_index = 0
        while col_index < len(boundaries) and x0 >= boundaries[col_index]:
            col_index += 1
        cells[col_index].append(str(word.get("text", "")).strip())

    return [" ".join(parts).strip() for parts in cells]


def _split_known_pdf_header_line(line: str) -> list[str]:
    normalized = re.sub(r"\s+", " ", line.strip()).lower()
    normalized = normalized.replace("netweight", "net weight").replace("grossweight", "gross weight")

    invoice_headers = [
        "marks & nos.",
        "marks & nos",
        "description of goods",
        "quantity",
        "unit price",
        "amount",
    ]
    packing_headers = [
        "packing no.",
        "packing no",
        "description of goods",
        "quantity",
        "net weight",
        "gross weight",
        "measurement",
    ]

    if "description of goods" not in normalized:
        return []

    if "unit price" in normalized and "amount" in normalized:
        return _headers_in_line(normalized, invoice_headers)
    if "packing" in normalized or "net weight" in normalized or "gross weight" in normalized:
        return _headers_in_line(normalized, packing_headers)
    return []


def _headers_in_line(normalized_line: str, headers: list[str]) -> list[str]:
    found: list[tuple[int, str]] = []
    for header in headers:
        pos = normalized_line.find(header)
        if pos >= 0:
            canonical = " ".join(part.capitalize() if part not in {"&"} else part for part in header.split())
            canonical = canonical.replace("Nos.", "Nos.").replace("No.", "No.")
            found.append((pos, canonical))
    found.sort(key=lambda item: item[0])

    result: list[str] = []
    seen: set[str] = set()
    for _, header in found:
        key = re.sub(r"[^a-z0-9]+", "", header.lower())
        if key not in seen:
            result.append(header)
            seen.add(key)
    return result


def _normalize_text_cell(value: object) -> str:
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return " ".join(text.split()).strip()


def _normalize_token(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", _normalize_text_cell(value).lower())


def _pad_rows(rows: list[list[str]]) -> list[list[str]]:
    if not rows:
        return []
    max_len = max(len(row) for row in rows)
    return [row + [""] * (max_len - len(row)) for row in rows]
