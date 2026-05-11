from __future__ import annotations

from io import BytesIO

import pandas as pd


def ocr_pdf_to_dataframes(
    file_name: str,
    data: bytes,
    *,
    scale: float = 1.25,
    min_score: float = 0.35,
) -> list[tuple[str, pd.DataFrame]]:
    try:
        import numpy as np
        import pypdfium2 as pdfium
        from rapidocr import RapidOCR
    except ModuleNotFoundError as exc:  # pragma: no cover - optional OCR stack
        raise RuntimeError("請先安裝 rapidocr、onnxruntime 與 pypdfium2，才能讀取掃描 PDF。") from exc

    document = pdfium.PdfDocument(BytesIO(data))
    engine = RapidOCR()
    tables: list[tuple[str, pd.DataFrame]] = []

    for page_index in range(len(document)):
        page = document[page_index]
        image = page.render(scale=scale).to_pil().convert("RGB")
        result = engine(np.array(image))
        rows = _ocr_result_to_rows(result, min_score=min_score)
        if rows:
            tables.append((f"PDF page {page_index + 1} OCR", pd.DataFrame(_pad_rows(rows))))

    return tables


def _ocr_result_to_rows(result: object, *, min_score: float) -> list[list[str]]:
    boxes = getattr(result, "boxes", None)
    txts = getattr(result, "txts", None)
    scores = getattr(result, "scores", None)
    if boxes is None or txts is None:
        return []

    items: list[dict[str, object]] = []
    for index, text in enumerate(txts):
        text = _normalize_text(text)
        if not text:
            continue
        score = float(scores[index]) if scores is not None and index < len(scores) else 1.0
        if score < min_score:
            continue

        box = boxes[index]
        xs = [float(point[0]) for point in box]
        ys = [float(point[1]) for point in box]
        items.append(
            {
                "text": text,
                "x": min(xs),
                "y": (min(ys) + max(ys)) / 2,
                "height": max(ys) - min(ys),
            }
        )

    if not items:
        return []

    median_height = sorted(float(item["height"]) for item in items)[len(items) // 2]
    y_threshold = max(8.0, median_height * 0.8)
    rows: list[list[dict[str, object]]] = []

    for item in sorted(items, key=lambda item: (float(item["y"]), float(item["x"]))):
        if not rows:
            rows.append([item])
            continue
        current_y = sum(float(row_item["y"]) for row_item in rows[-1]) / len(rows[-1])
        if abs(float(item["y"]) - current_y) <= y_threshold:
            rows[-1].append(item)
        else:
            rows.append([item])

    text_rows: list[list[str]] = []
    for row in rows:
        sorted_row = sorted(row, key=lambda item: float(item["x"]))
        text_rows.append([str(item["text"]) for item in sorted_row])

    return text_rows


def _pad_rows(rows: list[list[object]]) -> list[list[object]]:
    if not rows:
        return []
    max_len = max(len(row) for row in rows)
    return [row + [""] * (max_len - len(row)) for row in rows]


def _normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return " ".join(text.split()).strip()
