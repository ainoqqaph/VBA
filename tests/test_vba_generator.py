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

        self.assertIn("' 【出貨報單自動化處理系統】", vba_code)
        self.assertIn("Private Info As Variant", vba_code)
        self.assertRegex(vba_code, r"(?im)^\s*Sub\s+Main\s*\(")
        self.assertRegex(vba_code, r"(?im)^\s*Private\s+Sub\s+TINV\s*\(")
        self.assertRegex(vba_code, r"(?im)^\s*Private\s+Sub\s+TPKG\s*\(")
        self.assertRegex(vba_code, r"(?im)^\s*Private\s+Sub\s+CollectDataTinv\s*\(")
        self.assertRegex(vba_code, r"(?im)^\s*Private\s+Sub\s+WriteDataTinv\s*\(")
        self.assertRegex(vba_code, r"(?im)^\s*Private\s+Sub\s+CollectDataTpkg\s*\(")
        self.assertRegex(vba_code, r"(?im)^\s*Private\s+Sub\s+WriteDataTpkg\s*\(")
        self.assertRegex(vba_code, r"(?im)^\s*Private\s+Sub\s+InvNo回寫\s*\(")
        self.assertRegex(vba_code, r"(?im)^\s*Private\s+Sub\s+SayErrorCount\s*\(")

    def test_generated_vba_does_not_emit_streamlit_compatibility_procedures(self) -> None:
        vba_code = generate_op_temp_array_vba(
            [
                _rule("TINV", "Inv", "Tinv"),
                _rule("TPKG", "Pkg", "Tpkg"),
            ],
            menu_items=["Default", "QC", "阿格西"],
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

    def test_generated_vba_guards_carton_numbers_from_excel_date_conversion(self) -> None:
        vba_code = generate_op_temp_array_vba(
            [
                _rule("TINV", "Inv", "Tinv"),
                _rule("TPKG", "Pkg", "Tpkg"),
            ]
        )

        self.assertIn('Tst.Columns(1).NumberFormat = "@"', vba_code)
        self.assertIn("SafeExcelText(Write_CNO)", vba_code)
        self.assertIn("Private Function NormalizeDateLikeCartonText", vba_code)
        self.assertIn("Private Function MonthNameToNumber", vba_code)

    def test_generated_vba_procedure_list_excludes_old_extra_entrypoints(self) -> None:
        vba_code = generate_op_temp_array_vba(
            [
                _rule("TINV", "Inv", "Tinv"),
                _rule("TPKG", "Pkg", "Tpkg"),
            ]
        )

        procedure_names = set(
            re.findall(
                r"(?im)^\s*(?:Public|Private)?\s*(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*|InvNo回寫)\b",
                vba_code,
            )
        )
        forbidden_names = {
            "MainProcess",
            "MainProcess_Streamlit",
            "Streamlit_Main",
            "Streamlit_BuildOPTempArrays",
            "Streamlit_ExportOPTempArrayPreview",
            "Streamlit_MenuItems",
            "Streamlit_BindLegacyButtons",
            "Streamlit_TINV",
            "Streamlit_TPKG",
        }

        self.assertTrue(procedure_names.isdisjoint(forbidden_names), procedure_names & forbidden_names)


if __name__ == "__main__":
    unittest.main()
