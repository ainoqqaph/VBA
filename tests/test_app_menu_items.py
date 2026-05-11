from __future__ import annotations

from pathlib import Path
import unittest

from app import DEFAULT_SECTIONS, OP_TEMP_ARRAY_RUNTIME_RULES, _legacy_menu_items_from_classification


class LegacyMenuItemTests(unittest.TestCase):
    def test_rule_category_can_define_old_customer_menu_items(self) -> None:
        self.assertEqual(
            _legacy_menu_items_from_classification("Default, QC, \u963f\u683c\u897f", ""),
            ["Default", "QC", "\u963f\u683c\u897f"],
        )

    def test_single_customer_category_keeps_default_as_fallback(self) -> None:
        self.assertEqual(
            _legacy_menu_items_from_classification("\u963f\u683c\u897f", ""),
            ["Default", "\u963f\u683c\u897f"],
        )

    def test_empty_category_falls_back_to_default(self) -> None:
        self.assertEqual(_legacy_menu_items_from_classification("", ""), ["Default"])

    def test_tinv_fixed_title_defaults_empty_for_manual_entry(self) -> None:
        self.assertEqual(DEFAULT_SECTIONS["TINV"]["fixed_title"], "")

    def test_weight_mode_controls_are_hidden_from_streamlit_ui(self) -> None:
        app_source = Path("app.py").read_text(encoding="utf-8")

        self.assertNotIn("#### NW / GW 重量推算", app_source)
        self.assertNotIn('key=f"{key}_nw_mode"', app_source)
        self.assertNotIn('key=f"{key}_gw_mode"', app_source)

    def test_op_runtime_defaults_match_yi_sheng_final_layout(self) -> None:
        tinv = OP_TEMP_ARRAY_RUNTIME_RULES["TINV"]
        tpkg = OP_TEMP_ARRAY_RUNTIME_RULES["TPKG"]

        self.assertEqual(tinv["header_row"], 10)
        self.assertEqual(tinv["data_start_row"], 12)
        self.assertIn(("Unit Price", 13), tinv["mappings"])
        self.assertIn(("Amount", 14), tinv["mappings"])
        self.assertIn(("CTN", 6), tpkg["mappings"])
        self.assertIn(("Net Weight", 15), tpkg["mappings"])
        self.assertIn(("Gross Weight", 16), tpkg["mappings"])


if __name__ == "__main__":
    unittest.main()
