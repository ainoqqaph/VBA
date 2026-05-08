from __future__ import annotations

import traceback

import streamlit as st

from invoice_packing_cleaner.extractors import (
    MissingDependency,
    ParsedTable,
    UnsupportedFileType,
    parse_uploaded_file,
)
from invoice_packing_cleaner.profile_tools import (
    build_profile,
    dump_profile,
    load_profile,
    mappings_from_profile,
    target_columns_from_profile,
)
from invoice_packing_cleaner.table_tools import (
    DEFAULT_TARGET_COLUMNS,
    SourceColumn,
    build_source_options,
    find_default_source_index,
    prepare_structured_table,
    split_target_columns,
)
from invoice_packing_cleaner.template_tools import (
    TemplateCandidate,
    TemplateColumn,
    build_template_preview,
    parse_output_template_file,
)
from invoice_packing_cleaner.vba_generator import FieldMapping, generate_vba


st.set_page_config(page_title="Invoice / Packing List VBA Generator", layout="wide")


def main() -> None:
    st.title("Invoice / Packing List VBA 產生器")
    st.caption("用 OP 的最終輸出範本學欄位位置，再把客戶原始檔轉成固定格式 VBA。")

    imported_profile = _load_profile_from_sidebar()
    saved_mappings = mappings_from_profile(imported_profile)

    with st.sidebar:
        st.header("基本設定")
        customer_name = st.text_input(
            "客戶 / 規則名稱",
            value=str(imported_profile.get("customer_name", "")),
            placeholder="例如 TTI、客戶A、2026新版格式",
        )
        document_mode = st.radio(
            "客戶檔案型態",
            ("Invoice + Packing List 放一起", "Invoice / Packing List 分開給"),
            index=0,
        )

        source_profile = imported_profile.get("source", {})
        output_profile = imported_profile.get("output", {})
        lookup_default = 0 if source_profile.get("lookup_mode", "header") == "header" else 1
        lookup_mode_label = st.radio(
            "VBA 尋找來源欄位方式",
            ("依欄位名稱優先，找不到再用欄位位置", "只依欄位位置"),
            index=lookup_default,
            help="欄位會左右移動時建議用欄位名稱；表頭常重複或不穩定時可改用欄位位置。",
        )
        lookup_mode = "header" if lookup_mode_label.startswith("依欄位名稱") else "position"

        output_sheet_name = st.text_input(
            "輸出工作表名稱",
            value=str(output_profile.get("sheet_name", "CLEANED_INVOICE_PL")),
        )
        fixed_title = st.text_input(
            "固定補入的大標題",
            value=str(output_profile.get("fixed_title", "HAND TOOL")),
            help="原始資料沒有分類標題時，VBA 會主動寫入。若不需要可留空。",
        )

    target_columns = _target_template_section(imported_profile)

    uploaded_files = st.file_uploader(
        "上傳客戶原始 Invoice / Packing List 檔案",
        type=["xlsx", "xls", "xlsm", "csv", "pdf", "docx", "doc"],
        accept_multiple_files=True,
        help="PDF 若是掃描圖片，這版先不做 OCR；舊版 .doc 請先轉成 .docx。",
        key="source_files",
    )

    if document_mode == "Invoice / Packing List 分開給":
        st.info("分開給的檔案這版先用同一套欄位對應產生 VBA；之後可用規則 JSON 擴充成 Invoice 與 Packing List 各自一套對應。")

    if not uploaded_files:
        st.info("先上傳客戶原始檔，工具會顯示可用的 sheet、PDF 頁面或 Word 表格。")
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

    st.subheader("2. 指定來源表頭與資料列")
    max_rows = max(len(raw_df), 1)
    source_profile = imported_profile.get("source", {})
    default_header_row = _bounded_int(source_profile.get("header_row", 1), 1, max_rows)
    default_data_start_row = _bounded_int(source_profile.get("data_start_row", 2), 1, max_rows)

    setup_cols = st.columns(3)
    with setup_cols[0]:
        header_row = st.number_input("來源表頭列", min_value=1, max_value=max_rows, value=default_header_row, step=1)
    with setup_cols[1]:
        data_start_row = st.number_input(
            "來源資料開始列",
            min_value=1,
            max_value=max_rows,
            value=default_data_start_row,
            step=1,
        )
    with setup_cols[2]:
        st.metric("目前來源", selected_table.kind)

    structured_df, source_headers = prepare_structured_table(raw_df, int(header_row), int(data_start_row))
    source_options = build_source_options(source_headers)

    st.caption("下方預覽會用你指定的來源表頭列作為欄位名稱。")
    st.dataframe(structured_df.head(40), use_container_width=True)

    if not target_columns:
        st.error("請先設定或上傳最終輸出欄位。")
        return

    st.subheader("3. 對應來源欄位到最終格式")
    st.caption("左邊是 OP 最終格式偵測到的位置；右邊選客戶原始檔的來源欄位。未對應的欄位會輸出空白。")
    mappings = _mapping_editor(target_columns, source_options, saved_mappings)

    output_profile = imported_profile.get("output", {})
    default_output_header_row = _bounded_int(
        output_profile.get("header_row", st.session_state.get("detected_output_header_row", 1)),
        1,
        500,
    )
    default_output_data_start_row = _bounded_int(
        output_profile.get("data_start_row", st.session_state.get("detected_output_data_start_row", 2)),
        1,
        500,
    )

    st.subheader("4. 輸出位置設定")
    output_cols = st.columns(2)
    with output_cols[0]:
        output_header_row = st.number_input(
            "最終格式表頭列",
            min_value=1,
            max_value=500,
            value=default_output_header_row,
            step=1,
            help="通常等於 OP 範本裡欄位名稱所在的列。",
        )
    with output_cols[1]:
        output_data_start_row = st.number_input(
            "最終格式資料開始列",
            min_value=1,
            max_value=500,
            value=default_output_data_start_row,
            step=1,
            help="VBA 會從這一列開始寫入轉換後資料。",
        )

    st.subheader("5. 產生 VBA 與客戶規則")
    vba_code = generate_vba(
        mappings=mappings,
        header_row=int(header_row),
        data_start_row=int(data_start_row),
        output_sheet_name=output_sheet_name.strip() or "CLEANED_INVOICE_PL",
        output_header_row=int(output_header_row),
        output_data_start_row=int(output_data_start_row),
        lookup_mode=lookup_mode,
        fixed_title=fixed_title.strip(),
    )

    profile = build_profile(
        customer_name=customer_name.strip(),
        document_mode=document_mode,
        header_row=int(header_row),
        data_start_row=int(data_start_row),
        output_sheet_name=output_sheet_name.strip() or "CLEANED_INVOICE_PL",
        output_header_row=int(output_header_row),
        output_data_start_row=int(output_data_start_row),
        lookup_mode=lookup_mode,
        fixed_title=fixed_title.strip(),
        mappings=mappings,
    )
    profile_json = dump_profile(profile)
    safe_customer_name = _safe_file_stem(customer_name or "customer_rule")

    actions = st.columns(2)
    with actions[0]:
        st.download_button(
            "下載 VBA 模組 .bas",
            data=vba_code.encode("utf-8-sig"),
            file_name=f"{safe_customer_name}_CleanInvoicePackingList.bas",
            mime="text/plain",
        )
    with actions[1]:
        st.download_button(
            "下載客戶規則 JSON",
            data=profile_json.encode("utf-8-sig"),
            file_name=f"{safe_customer_name}_rule.json",
            mime="application/json",
            help="下次遇到同客戶或同格式，直接匯入這個 JSON，不必重新對欄位。",
        )

    st.code(vba_code, language="vbnet")


def _load_profile_from_sidebar() -> dict:
    with st.sidebar:
        st.header("客戶規則")
        profile_file = st.file_uploader(
            "匯入已儲存規則 JSON",
            type=["json"],
            key="profile_json",
            help="同客戶下次直接匯入規則，不用重新找表頭與對欄位。",
        )

    if not profile_file:
        return {}

    try:
        profile = load_profile(profile_file.getvalue())
        st.sidebar.success("已載入規則 JSON。")
        return profile
    except Exception as exc:
        st.sidebar.error(f"規則 JSON 讀取失敗：{exc}")
        return {}


def _target_template_section(imported_profile: dict) -> list[TemplateColumn]:
    st.subheader("A. OP 最終輸出格式範本")
    st.caption("上傳 OP 以前手打好的最終格式檔，系統會偵測欄位名稱與輸出位置。這是用來學格式，不是客戶原始檔。")

    profile_columns = target_columns_from_profile(imported_profile)
    template_file = st.file_uploader(
        "上傳 OP 最終格式範本（Excel / CSV，可選）",
        type=["xlsx", "xls", "xlsm", "csv"],
        key="output_template",
    )

    if template_file:
        try:
            candidates = parse_output_template_file(template_file.name, template_file.getvalue())
        except Exception as exc:
            st.error(f"最終格式範本讀取失敗：{exc}")
            candidates = []

        if candidates:
            selected_candidate = _select_template_candidate(candidates)
            st.session_state["detected_output_header_row"] = selected_candidate.header_row
            st.session_state["detected_output_data_start_row"] = selected_candidate.data_start_row

            preview_cols = st.columns([1, 2])
            with preview_cols[0]:
                st.markdown("偵測到的輸出欄位位置")
                st.dataframe(
                    build_template_preview(selected_candidate.columns, selected_candidate.header_row),
                    use_container_width=True,
                    hide_index=True,
                )
            with preview_cols[1]:
                st.markdown("範本預覽")
                st.dataframe(selected_candidate.dataframe.head(40), use_container_width=True)

            return selected_candidate.columns

        st.warning("沒有自動偵測到像表頭的列。你可以改用下方手動欄位清單。")

    with st.expander("沒有範本時：手動輸入最終欄位", expanded=not profile_columns):
        default_columns = profile_columns or [
            TemplateColumn(name, index)
            for index, name in enumerate(DEFAULT_TARGET_COLUMNS, start=1)
        ]
        target_columns_text = st.text_area(
            "最終輸出欄位（一行一個，也可用逗號分隔）",
            value="\n".join(column.name for column in default_columns),
            height=220,
        )
        names = split_target_columns(target_columns_text)
        return [TemplateColumn(name, index) for index, name in enumerate(names, start=1)]

    return profile_columns


def _select_template_candidate(candidates: list[TemplateCandidate]) -> TemplateCandidate:
    labels = [candidate.label for candidate in candidates]
    selected_label = st.selectbox("選擇偵測到的最終格式表頭列", labels)
    return candidates[labels.index(selected_label)]


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
    selected_label = st.selectbox("選擇要建立規則的來源表格 / 頁面", labels)
    return parsed_tables[labels.index(selected_label)]


def _mapping_editor(
    target_columns: list[TemplateColumn],
    source_options: list[SourceColumn],
    saved_mappings: dict[str, dict],
) -> list[FieldMapping]:
    option_labels = [option.label for option in source_options]
    mappings: list[FieldMapping] = []

    for row_start in range(0, len(target_columns), 2):
        cols = st.columns(2)
        for offset, column in enumerate(cols):
            target_index = row_start + offset
            if target_index >= len(target_columns):
                continue

            target = target_columns[target_index]
            saved_mapping = saved_mappings.get(target.name, {})
            default_index = _default_source_option_index(target.name, source_options, saved_mapping)

            with column:
                st.caption(f"輸出位置：{target.column_letter} 欄")
                selected_label = st.selectbox(
                    target.name,
                    option_labels,
                    index=default_index,
                    key=f"map_{target_index}_{target.name}_{target.column_index}",
                )
                source = source_options[option_labels.index(selected_label)]
                mappings.append(
                    FieldMapping(
                        target=target.name,
                        source_header=source.header,
                        source_index=source.index,
                        target_col=target.column_index,
                    )
                )

    return mappings


def _default_source_option_index(
    target: str,
    source_options: list[SourceColumn],
    saved_mapping: dict,
) -> int:
    saved_index = int(saved_mapping.get("source_index") or 0)
    if saved_index > 0:
        for option_position, option in enumerate(source_options):
            if option.index == saved_index:
                return option_position

    saved_header = str(saved_mapping.get("source_header", "")).strip()
    if saved_header:
        for option_position, option in enumerate(source_options):
            if option.header == saved_header:
                return option_position

    return find_default_source_index(target, source_options)


def _show_next_steps() -> None:
    with st.expander("建議工作流程"):
        st.markdown(
            "1. 上傳 OP 以前手打好的最終輸出格式，讓系統偵測輸出欄位位置。\n"
            "2. 上傳客戶原始 invoice / packing list。\n"
            "3. 選擇正確的 sheet、PDF 頁面或 Word 表格。\n"
            "4. 指定來源表頭列與來源資料開始列。\n"
            "5. 把來源欄位對應到 OP 最終格式欄位。\n"
            "6. 下載 VBA 和客戶規則 JSON。"
        )


def _bounded_int(value: object, minimum: int, maximum: int) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        number = minimum
    return max(minimum, min(maximum, number))


def _safe_file_stem(value: str) -> str:
    cleaned = "".join(char if char.isalnum() or char in "-_" else "_" for char in value.strip())
    return cleaned.strip("_") or "customer_rule"


if __name__ == "__main__":
    main()
