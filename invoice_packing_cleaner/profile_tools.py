from __future__ import annotations

import json
from typing import Any

from invoice_packing_cleaner.template_tools import TemplateColumn
from invoice_packing_cleaner.vba_generator import FieldMapping


PROFILE_VERSION = 1


def load_profile(data: bytes) -> dict[str, Any]:
    text = data.decode("utf-8-sig")
    profile = json.loads(text)
    if not isinstance(profile, dict):
        raise ValueError("規則檔格式錯誤。")
    return profile


def dump_profile(profile: dict[str, Any]) -> str:
    profile = {"version": PROFILE_VERSION, **profile}
    return json.dumps(profile, ensure_ascii=False, indent=2)


def mappings_from_profile(profile: dict[str, Any]) -> dict[str, dict[str, Any]]:
    mappings = profile.get("mappings", [])
    if not isinstance(mappings, list):
        return {}
    result: dict[str, dict[str, Any]] = {}
    for mapping in mappings:
        if isinstance(mapping, dict) and mapping.get("target"):
            result[str(mapping["target"])] = mapping
    return result


def target_columns_from_profile(profile: dict[str, Any]) -> list[TemplateColumn]:
    mappings = profile.get("mappings", [])
    columns: list[TemplateColumn] = []
    if not isinstance(mappings, list):
        return columns

    for index, mapping in enumerate(mappings, start=1):
        if not isinstance(mapping, dict):
            continue
        target = str(mapping.get("target", "")).strip()
        if not target:
            continue
        target_col = int(mapping.get("target_col") or index)
        columns.append(TemplateColumn(target, target_col))

    return columns


def build_profile(
    customer_name: str,
    document_mode: str,
    header_row: int,
    data_start_row: int,
    output_sheet_name: str,
    output_header_row: int,
    output_data_start_row: int,
    lookup_mode: str,
    fixed_title: str,
    mappings: list[FieldMapping],
) -> dict[str, Any]:
    return {
        "customer_name": customer_name,
        "document_mode": document_mode,
        "source": {
            "header_row": header_row,
            "data_start_row": data_start_row,
            "lookup_mode": lookup_mode,
        },
        "output": {
            "sheet_name": output_sheet_name,
            "header_row": output_header_row,
            "data_start_row": output_data_start_row,
            "fixed_title": fixed_title,
        },
        "mappings": [
            {
                "target": mapping.target,
                "target_col": mapping.target_col,
                "source_header": mapping.source_header,
                "source_index": mapping.source_index,
            }
            for mapping in mappings
        ],
    }
