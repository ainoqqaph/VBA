from __future__ import annotations

from invoice_packing_cleaner.pdf_text_tools import words_to_rows


def _word(text: str, top: float, x0: float) -> dict[str, float | str]:
    return {"text": text, "top": top, "x0": x0}


def test_words_to_rows_preserves_packing_columns() -> None:
    words = [
        _word("Packing", 10, 27),
        _word("No.", 10, 69),
        _word("Description", 10, 139),
        _word("of", 10, 199),
        _word("Goods", 10, 211),
        _word("Quantity", 10, 306),
        _word("NetWe", 10, 365),
        _word("ight", 10, 400),
        _word("GrossWeight", 10, 434),
        _word("Me", 10, 504),
        _word("asurement", 10, 521),
        _word("1-4", 24, 18),
        _word("ACETex", 24, 93),
        _word("GT150-I", 24, 134),
        _word("@1,470.00MTK", 24, 277),
        _word("@568.38KGS", 24, 371),
        _word("@613.38KGS", 24, 454),
    ]

    rows = words_to_rows(words)

    assert rows[0] == [
        "Packing No.",
        "Description Of Goods",
        "Quantity",
        "Net Weight",
        "Gross Weight",
        "Measurement",
    ]
    assert rows[1] == [
        "1-4",
        "ACETex GT150-I",
        "@1,470.00MTK",
        "@568.38KGS",
        "@613.38KGS",
        "",
    ]
