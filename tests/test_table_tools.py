from __future__ import annotations

import unittest

from invoice_packing_cleaner.table_tools import build_source_options


class SourceOptionLabelTests(unittest.TestCase):
    def test_source_options_use_preview_column_numbers(self) -> None:
        options = build_source_options(["MARKS & NOS.", "DESCRIPTION", "Column C", "QUANTITY"])

        self.assertEqual(
            [option.label for option in options],
            ["\uff08\u7559\u7a7a\uff09", "0 | MARKS & NOS.", "1 | DESCRIPTION", "2", "3 | QUANTITY"],
        )
        self.assertEqual(options[2].index, 2)
        self.assertEqual(options[2].header, "DESCRIPTION")


if __name__ == "__main__":
    unittest.main()
