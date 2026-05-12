from __future__ import annotations

import re
import unittest

from invoice_packing_cleaner.vba_generator import (
    FieldMapping,
    SheetTransformRule,
    generate_op_temp_array_vba,
)


def _rule(section: str, source_sheet: str, output_sheet: str) -> SheetTransformRule:
    return SheetTransformRule(
        procedure_name=f"Clean{section}",
        source_sheet_name=source_sheet,
        output_sheet_name=output_sheet,
        mappings=[
            FieldMapping(target="Item No", source_header="Item No", source_index=1, target_col=1),
            FieldMapping(target="Description", source_header="Description", source_index=2, target_col=2),
            FieldMapping(target="Quantity", source_header="Quantity", source_index=3, target_col=3),
        ],
        header_row=1,
        data_start_row=2,
        output_header_row=1,
        output_data_start_row=2,
        fixed_title="",
    )


class OpTempArrayVbaGeneratorTests(unittest.TestCase):
    def test_generated_vba_uses_strict_op_module_shape(self) -> None:
        vba_code = generate_op_temp_array_vba(
            [
                _rule("TINV", "Inv", "Tinv"),
                _rule("TPKG", "Pkg", "Tpkg"),
            ]
        )

        self.assertIn("' \u3010\u51fa\u8ca8\u5831\u55ae\u81ea\u52d5\u5316\u8655\u7406\u7cfb\u7d71\u3011", vba_code)
        self.assertIn("Private Info As Variant", vba_code)
        expected_names = [
            "Main",
            "TINV",
            "CollectDataTinv",
            "WriteDataTinv",
            "TPKG",
            "CollectDataTpkg",
            "WriteDataTpkg",
            "InvNo\u56de\u5beb",
            "SayErrorCount",
            "GetActualTotals",
            "GetSafeValue",
            "GetInnerValue",
            "JoinPO",
            "GetUsedQty",
            "UpdateUsedQty",
            "HasEnoughQty",
            "IsFullyUsed",
        ]
        procedure_names = re.findall(
            r"(?im)^\s*(?:Public|Private)?\s*(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*|InvNo\u56de\u5beb)\b",
            vba_code,
        )
        self.assertEqual(expected_names, procedure_names)

        forbidden = [
            "Private INVcollection",
            "Private PKGcollection",
            "Private Const ENABLE_MULTI_BOX_DETAIL",
            "Dim foundHeader",
            "Set foundHeader",
            "Private Function LastUsedRow",
            "Private Function ResolveOpColumn",
            "Private Sub WriteHeaders",
            "Private Function InCollection",
            "Private Function CellTextAt",
            "Private Function SafeExcelText",
            "Private Function BuildCheckText",
            "CellTextAt(",
            "SafeExcelText(",
            "BuildCheckText(",
        ]
        for text in forbidden:
            self.assertNotIn(text, vba_code)

    def test_generated_vba_does_not_emit_streamlit_compatibility_procedures(self) -> None:
        vba_code = generate_op_temp_array_vba(
            [
                _rule("TINV", "Inv", "Tinv"),
                _rule("TPKG", "Pkg", "Tpkg"),
            ],
            menu_items=["Default", "QC", "\u963f\u683c\u897f"],
        )

        forbidden = [
            "Streamlit_",
            "MainProcess_Streamlit",
            "Auto_Open",
            "ExportOPTempArrayPreview",
            "Debug_INVcollection",
            "Debug_PKGcollection",
            "BindLegacyButtons",
        ]
        for text in forbidden:
            self.assertNotIn(text, vba_code)

    def test_generated_vba_writes_hand_tool_when_user_sets_fixed_title(self) -> None:
        tinv_rule = _rule("TINV", "Inv", "Tinv")
        custom_tinv_rule = SheetTransformRule(
            procedure_name=tinv_rule.procedure_name,
            source_sheet_name=tinv_rule.source_sheet_name,
            output_sheet_name=tinv_rule.output_sheet_name,
            mappings=tinv_rule.mappings,
            header_row=tinv_rule.header_row,
            data_start_row=tinv_rule.data_start_row,
            output_header_row=tinv_rule.output_header_row,
            output_data_start_row=tinv_rule.output_data_start_row,
            fixed_title="HAND TOOL",
        )
        vba_code = generate_op_temp_array_vba(
            [
                custom_tinv_rule,
                _rule("TPKG", "Pkg", "Tpkg"),
            ]
        )

        self.assertIn('Const FIXED_TITLE As String = "HAND TOOL"', vba_code)
        self.assertNotIn("HAND TOOLS", vba_code)

    def test_generated_vba_keeps_carton_numbers_as_text_without_extra_helpers(self) -> None:
        vba_code = generate_op_temp_array_vba(
            [
                _rule("TINV", "Inv", "Tinv"),
                _rule("TPKG", "Pkg", "Tpkg"),
            ]
        )

        self.assertIn('Tst.Columns(1).NumberFormat = "@"', vba_code)
        self.assertIn('.Cells(newRow, 1).Value = Write_CNO', vba_code)
        self.assertNotIn("Private Function NormalizeDateLikeCartonText", vba_code)
        self.assertNotIn("Private Function MonthNameToNumber", vba_code)

    def test_generated_vba_uses_legacy_getcolumn_flow_when_saved_mapping_is_blank(self) -> None:
        empty_tinv = SheetTransformRule(
            procedure_name="CleanTINV",
            source_sheet_name="Inv",
            output_sheet_name="Tinv",
            mappings=[
                FieldMapping(target="Description", source_header="", source_index=0, target_col=1),
                FieldMapping(target="Quantity", source_header="", source_index=0, target_col=2),
                FieldMapping(target="Unit Price", source_header="", source_index=0, target_col=3),
                FieldMapping(target="Amount", source_header="", source_index=0, target_col=4),
            ],
            header_row=2,
            data_start_row=3,
            output_header_row=1,
            output_data_start_row=2,
            fixed_title="",
        )
        empty_tpkg = SheetTransformRule(
            procedure_name="CleanTPKG",
            source_sheet_name="Pkg",
            output_sheet_name="Tpkg",
            mappings=[
                FieldMapping(target="CTN", source_header="", source_index=0, target_col=1),
                FieldMapping(target="Unit Qty", source_header="", source_index=0, target_col=2),
                FieldMapping(target="Gross Weight", source_header="", source_index=0, target_col=3),
            ],
            header_row=2,
            data_start_row=3,
            output_header_row=1,
            output_data_start_row=2,
            fixed_title="",
        )

        vba_code = generate_op_temp_array_vba([empty_tinv, empty_tpkg], lookup_mode="position")

        self.assertIn("Set map = ExploreFieldsWithSmartMatch", vba_code)
        self.assertIn('colDSC = GetColumn(Dst, map, "\u54c1\u9805\u63cf\u8ff0", "*DESC.*")', vba_code)
        self.assertIn('colQty = GetColumn(Dst, map, "\u6578\u91cf", "*QTY*")', vba_code)
        self.assertNotIn("Const fallbackColDSC", vba_code)
        self.assertNotIn("ResolveOpColumn", vba_code)

    def test_get_actual_totals_sits_before_bottom_helper_area(self) -> None:
        vba_code = generate_op_temp_array_vba(
            [
                _rule("TINV", "Inv", "Tinv"),
                _rule("TPKG", "Pkg", "Tpkg"),
            ]
        )

        self.assertLess(vba_code.index("Public Sub GetActualTotals"), vba_code.index("Function GetSafeValue"))
        self.assertLess(vba_code.index("Private Sub SayErrorCount"), vba_code.index("Public Sub GetActualTotals"))
        self.assertLess(vba_code.index("\u4ee5\u4e0b\u7686\u70ba\u7cfb\u7d71"), vba_code.index("Function GetSafeValue"))


if __name__ == "__main__":
    unittest.main()
