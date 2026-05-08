from __future__ import annotations

import traceback

import streamlit as st

from invoice_packing_cleaner.extractors import (
    MissingDependency,
    ParsedTable,
    UnsupportedFileType,
    parse_uploaded_file,
)
from invoice_packing_cleaner.table_tools import (
    DEFAULT_TARGET_COLUMNS,
    SourceColumn,
    build_source_options,
    find_default_source_index,
    prepare_structured_table,
    split_target_columns,
)
from invoice_packing_cleaner.vba_generator import FieldMapping, generate_vba


st.set_page_config(page_title="Invoice / Packing List VBA Generator", layout="wide")


def main() -> None:
    st.title("Invoice / Packing List VBA 產生器")
    st.caption("上傳客戶檔案，確認資料區與欄位對應後，產生可貼到 Excel 的 VBA 模組。")

    with st.sidebar:
        st.header("設定")
        document_mode = st.radio(
            "客戶檔案型態",
            ("Invoice + Packing List 放一起", "Invoice / Packing List 分開給"),
        )
        lookup_mode_label = st.radio(
            "VBA 尋找來源欄位方式",
            ("依欄位名稱優先，找不到再用欄位位置", "只依欄位位置"),
            help="同一客戶欄位會移動時建議用欄位名稱；表頭常重複或不固定時建議用欄位位置。",
        )
        lookup_mode = "header" if lookup_mode_label.startswith("依欄位名稱") else "position"
        output_sheet_name = st.text_input("輸出工作表名稱", value="CLEANED_INVOICE_PL")
        fixed_title = st.text_input(
            "固定補入的大標題",
            value="HAND TOOL",
            help="原始資料沒有這個分類標題時，VBA 會主動寫到輸出資料最前面；不需要時可留空。",
        )
        target_columns_text = st.text_area(
            "最終輸出欄位（一行一個，也可用逗號分隔）",
            value="\n".join(DEFAULT_TARGET_COLUMNS),
            height=280,
        )

    uploaded_files = st.file_uploader(
        "上傳 Invoice / Packing List 檔案",
        type=["xlsx", "xls", "xlsm", "csv", "pdf", "docx", "doc"],
        accept_multiple_files=True,
        help="PDF 若是掃描圖片，這版先不做 OCR；舊版 .doc 請先轉成 .docx。",
    )

    if document_mode == "Invoice / Packing List 分開給":
        st.info("這版先用同一套欄位對應產生單一來源 VBA。你給我實際 VBA 範例後，我會把分開來源的合併規則補進來。")

    if not uploaded_files:
        st.info("先上傳一份 Excel、PDF 或 Word 檔，工具會顯示可用的表格或文字區塊。")
        _show_next_steps()
        return

    parsed_tables = _parse_uploaded_files(uploaded_files)
    if not parsed_tables:
        st.error("沒有讀到可用資料。")
        return

    selected_table = _select_table(parsed_tables)
    raw_df = selected_table.dataframe

    if selected_table.note:
        st.warning(selected_table.note)

    if raw_df.empty:
        st.error("這個區塊沒有資料可預覽。")
        return

    st.subheader("1. 原始資料預覽")
    st.dataframe(raw_df.head(80), use_container_width=True)

    st.subheader("2. 指定表頭與資料列")
    max_rows = max(len(raw_df), 1)
    setup_cols = st.columns(3)
    with setup_cols[0]:
        header_row = st.number_input("表頭列", min_value=1, max_value=max_rows, value=1, step=1)
    with setup_cols[1]:
        data_start_row = st.number_input(
            "資料開始列",
            min_value=1,
            max_value=max_rows,
            value=min(2, max_rows),
            step=1,
        )
    with setup_cols[2]:
        st.metric("目前來源", selected_table.kind)

    structured_df, source_headers = prepare_structured_table(raw_df, int(header_row), int(data_start_row))
    source_options = build_source_options(source_headers)

    st.caption("下方預覽會用你指定的表頭列作為欄位名稱。")
    st.dataframe(structured_df.head(40), use_container_width=True)

    target_columns = split_target_columns(target_columns_text)
    if not target_columns:
        st.error("請至少設定一個最終輸出欄位。")
        return

    st.subheader("3. 對應最終欄位")
    st.caption("每個最終欄位選擇客戶檔案中的來源欄位；未對應的欄位會在 VBA 輸出空白。")
    mappings = _mapping_editor(target_columns, source_options)

    st.subheader("4. 產生 VBA")
    vba_code = generate_vba(
        mappings=mappings,
        header_row=int(header_row),
        data_start_row=int(data_start_row),
        output_sheet_name=output_sheet_name.strip() or "CLEANED_INVOICE_PL",
        lookup_mode=lookup_mode,
        fixed_title=fixed_title.strip(),
    )

    st.download_button(
        "下載 VBA 模組 .bas",
        data=vba_code.encode("utf-8-sig"),
        file_name="CleanInvoicePackingList.bas",
        mime="text/plain",
    )
    st.code(vba_code, language="vbnet")

    with st.expander("之後要套進正式規則時，我需要的資料"):
        st.markdown(
            "- 你現在用的 VBA 範例\n"
            "- 最終 Excel 欄位順序與欄名\n"
            "- Invoice 與 Packing List 若分開時，要用哪個鍵值合併，例如料號、項次、箱號\n"
            "- 客戶常見特殊情況，例如合併儲存格、跨頁 PDF、重量在頁尾、幣別在表頭"
        )


def _parse_uploaded_files(uploaded_files) -> list[ParsedTable]:
    parsed_tables: list[ParsedTable] = []

    for uploaded_file in uploaded_files:
        try:
            parsed_tables.extend(parse_uploaded_file(uploaded_file.name, uploaded_file.getvalue()))
        except UnsupportedFileType as exc:
            st.warning(f"{uploaded_file.name}：{exc}")
        except MissingDependency as exc:
            st.error(str(exc))
        except Exception as exc:  # pragma: no cover - Streamlit error presentation
            st.error(f"{uploaded_file.name} 讀取失敗：{exc}")
            with st.expander(f"{uploaded_file.name} 錯誤細節"):
                st.code(traceback.format_exc())

    return parsed_tables


def _select_table(parsed_tables: list[ParsedTable]) -> ParsedTable:
    labels = [table.label for table in parsed_tables]
    selected_label = st.selectbox("選擇要建立規則的表格 / 頁面", labels)
    return parsed_tables[labels.index(selected_label)]


def _mapping_editor(target_columns: list[str], source_options: list[SourceColumn]) -> list[FieldMapping]:
    option_labels = [option.label for option in source_options]
    mappings: list[FieldMapping] = []

    for row_start in range(0, len(target_columns), 2):
        cols = st.columns(2)
        for offset, column in enumerate(cols):
            target_index = row_start + offset
            if target_index >= len(target_columns):
                continue

            target = target_columns[target_index]
            default_index = find_default_source_index(target, source_options)
            with column:
                selected_label = st.selectbox(
                    target,
                    option_labels,
                    index=default_index,
                    key=f"map_{target_index}_{target}",
                )
                source = source_options[option_labels.index(selected_label)]
                mappings.append(
                    FieldMapping(
                        target=target,
                        source_header=source.header,
                        source_index=source.index,
                    )
                )

    return mappings


def _show_next_steps() -> None:
    with st.expander("建議工作流程"):
        st.markdown(
            "1. 上傳客戶提供的 invoice / packing list。\n"
            "2. 選擇正確的 sheet、PDF 頁面或 Word 表格。\n"
            "3. 指定表頭列與資料開始列。\n"
            "4. 把客戶欄位對應到你的最終欄位。\n"
            "5. 下載 VBA，貼到 Excel VBA 編輯器執行。"
        )


if __name__ == "__main__":
    main()
