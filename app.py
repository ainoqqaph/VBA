from __future__ import annotations

import traceback
from dataclasses import dataclass
from typing import Any

import streamlit as st

from invoice_packing_cleaner.extractors import (
    MissingDependency,
    ParsedTable,
    UnsupportedFileType,
    parse_uploaded_file,
)
from invoice_packing_cleaner.profile_tools import dump_profile, load_profile
from invoice_packing_cleaner.table_tools import (
    DEFAULT_TARGET_COLUMNS,
    OP_TINV_COLUMNS,
    OP_TPKG_COLUMNS,
    SourceColumn,
    build_source_options,
    find_default_source_index,
    prepare_structured_table,
    split_target_columns,
    suggest_header_row,
)
from invoice_packing_cleaner.template_tools import (
    TemplateCandidate,
    TemplateColumn,
    build_template_preview,
    parse_output_template_file,
)
from invoice_packing_cleaner.vba_generator import (
    FieldMapping,
    SheetTransformRule,
    generate_op_temp_array_vba,
    generate_workbook_vba,
)


st.set_page_config(page_title="Invoice / Packing List VBA Generator", layout="wide")


DEFAULT_SECTIONS = {
    "TINV": {
        "label": "TINV / Invoice",
        "source_sheet": "Inv",
        "output_sheet": "Tinv",
        "procedure": "CleanTINV",
        "fixed_title": "HAND TOOL",
    },
    "TPKG": {
        "label": "TPKG / Packing List",
        "source_sheet": "Pkg",
        "output_sheet": "Tpkg",
        "procedure": "CleanTPKG",
        "fixed_title": "",
    },
}

WEIGHT_MODE_LABELS = {
    "source_is_unit": "來源是單箱/單件重量：直接放進 tempArray",
    "source_is_total": "來源是總重量：先除以箱數再放進 tempArray",
}


OP_TEMP_ARRAY_RUNTIME_RULES = {
    "TINV": {
        "header_row": 3,
        "data_start_row": 4,
        "mappings": [
            ("PO No.", 1),
            ("Item No.", 3),
            ("Line No.", 4),
            ("Description of Goods", 6),
            ("Quantity", 7),
            ("Unit", 0),
            ("Unit Price", 8),
            ("Amount", 9),
            ("Brand", 2),
            ("Customer Item", 3),
        ],
    },
    "TPKG": {
        "header_row": 3,
        "data_start_row": 4,
        "mappings": [
            ("Customer PO", 1),
            ("PO No.", 1),
            ("Item No.", 3),
            ("Description of Goods", 6),
            ("Quantity", 7),
            ("Unit Qty", 10),
            ("Unit", 0),
            ("CTN", 14),
            ("Net Weight", 11),
            ("Gross Weight", 12),
            ("Measurement", 13),
        ],
    },
}


@dataclass(frozen=True)
class OutputTemplateConfig:
    key: str
    columns: list[TemplateColumn]
    header_row: int
    data_start_row: int
    label: str


def _target_columns_for_section(
    key: str,
    config: OutputTemplateConfig,
    op_temp_array_mode: bool,
) -> list[TemplateColumn]:
    if not op_temp_array_mode:
        return config.columns

    names = OP_TINV_COLUMNS if key == "TINV" else OP_TPKG_COLUMNS
    return [TemplateColumn(name, index) for index, name in enumerate(names, start=1)]


def main() -> None:
    _inject_soft_theme_css()
    st.markdown(
        """
        <section class="app-hero">
            <p class="app-kicker">報關文件清洗工具</p>
            <h1>Invoice / Packing List VBA 產生器</h1>
            <p>把客戶原始 invoice、packing list 對應成 OP 可維護的 VBA 規則。</p>
        </section>
        """,
        unsafe_allow_html=True,
    )

    imported_profile = _load_profile_from_sidebar()

    with st.sidebar:
        st.header("基本設定")
        customer_name = st.text_input(
            "客戶 / 規則名稱",
            value=str(imported_profile.get("customer_name", "")),
            placeholder="例如 TTI、客戶A、2026新版格式",
        )
        imported_classification = _classification_profile(imported_profile)
        with st.expander("客戶規則分類", expanded=True):
            end_customer_name = st.text_input(
                "終端客戶 / 買方 / 品牌",
                value=str(imported_classification.get("end_customer_name", "")),
                placeholder="例如：客戶的客戶、品牌、Consignee",
            )
            rule_category = st.text_input(
                "規則分類",
                value=str(imported_classification.get("rule_category", "")),
                placeholder="例如：一般格式 / 總重推算 / 特殊品名",
            )
            rule_tags = st.text_input(
                "搜尋標籤",
                value=", ".join(imported_classification.get("tags", [])),
                placeholder="例如：GW總重, NW單箱, HAND TOOL",
            )
            rule_note = st.text_area(
                "規則備註",
                value=str(imported_classification.get("note", "")),
                height=90,
            )
        lookup_mode_label = st.radio(
            "VBA 尋找來源欄位方式",
            ("依欄位名稱優先，找不到再用欄位位置", "只依欄位位置"),
            index=0 if imported_profile.get("lookup_mode", "header") == "header" else 1,
            help="欄位會左右移動時建議用欄位名稱；表頭常重複或不穩定時可改用欄位位置。",
        )
        lookup_mode = "header" if lookup_mode_label.startswith("依欄位名稱") else "position"
        vba_output_mode = st.radio(
            "VBA 輸出模式",
            ("OP tempArray / Collection 格式", "直接輸出 Tinv/Tpkg 工作表"),
            index=0,
            help="OP 模式會產生 INVcollection / PKGcollection，資料會照既有 tempArray 順序打包。",
        )

    template_configs = _target_template_section(imported_profile)

    uploaded_files = st.file_uploader(
        "上傳客戶原始 Invoice / Packing List 檔案",
        type=["xlsx", "xls", "xlsm", "csv", "pdf", "docx", "doc"],
        accept_multiple_files=True,
        help="可以一次上傳 invoice 和 packing list；也可以上傳 invoice+packing list 放在同一檔的檔案。",
        key="source_files",
    )

    if not uploaded_files:
        st.info("先上傳客戶原始檔，工具會顯示可用的 sheet、PDF 頁面或 Word 表格。")
        _show_next_steps()
        return

    parsed_tables = _parse_uploaded_files(uploaded_files)
    if not parsed_tables:
        st.error("沒有讀到可用資料。")
        return

    with st.expander("已讀取到的來源分頁 / 表格"):
        st.write([table.label for table in parsed_tables])

    rules: list[SheetTransformRule] = []
    profile_sheets: dict[str, Any] = {}

    st.subheader("B. 分別設定 TINV 和 TPKG 來源對應")
    st.info("這裡有兩個分頁要設定：先完成 1. TINV / Invoice，再點右邊的 2. TPKG / Packing List。")
    tab_labels = ["1. TINV / Invoice", "2. TPKG / Packing List"]
    tabs = st.tabs(tab_labels)

    for tab, key, tab_label in zip(tabs, DEFAULT_SECTIONS, tab_labels):
        with tab:
            st.markdown(f"## {tab_label}")
            if key == "TINV":
                st.caption("這一頁設定 Invoice 來源，通常選 INVOICE 分頁，VBA 來源工作表預設是 Inv。")
            else:
                st.caption("這一頁設定 Packing List 來源，通常選 PACKING 分頁，VBA 來源工作表預設是 Pkg。")

            config = template_configs.get(key)
            if not config or not config.columns:
                st.warning(f"尚未設定 {key} 的最終格式欄位。請回到 A 區選擇 {key} 最終格式。")
                continue

            rule, sheet_profile = _section_workflow(
                key=key,
                config=config,
                parsed_tables=parsed_tables,
                imported_profile=imported_profile,
                op_temp_array_mode=vba_output_mode.startswith("OP tempArray"),
            )
            if rule:
                rules.append(rule)
                profile_sheets[key] = sheet_profile

    if not rules:
        st.error("請至少完成一個 TINV 或 TPKG 的欄位對應。")
        return

    st.subheader("C. 產生 VBA 與客戶規則")
    if vba_output_mode.startswith("OP tempArray"):
        vba_code = generate_op_temp_array_vba(rules, lookup_mode="position")
        vba_file_suffix = "OP_TempArrays"
    else:
        vba_code = generate_workbook_vba(rules, lookup_mode=lookup_mode)
        vba_file_suffix = "TINV_TPKG"

    profile = {
        "customer_name": customer_name.strip(),
        "classification": {
            "company_customer_name": customer_name.strip(),
            "end_customer_name": end_customer_name.strip(),
            "rule_category": rule_category.strip(),
            "tags": _split_tags(rule_tags),
            "note": rule_note.strip(),
        },
        "lookup_mode": lookup_mode,
        "vba_output_mode": vba_output_mode,
        "sheets": profile_sheets,
    }
    profile_json = dump_profile(profile)
    safe_customer_name = _safe_file_stem(
        "_".join(
            part
            for part in [customer_name, end_customer_name, rule_category]
            if part.strip()
        )
        or "customer_rule"
    )

    actions = st.columns(2)
    with actions[0]:
        st.download_button(
            "下載 VBA 模組 .bas",
            data=vba_code.encode("utf-8-sig"),
            file_name=f"{safe_customer_name}_{vba_file_suffix}.bas",
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


def _inject_soft_theme_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            --soft-bg: #fbf7f2;
            --soft-panel: #fffdf9;
            --soft-panel-strong: #ffffff;
            --soft-line: #eaded3;
            --soft-muted: #766d64;
            --soft-text: #302c28;
            --soft-accent: #b96f5a;
            --soft-accent-dark: #8f4f3f;
            --soft-sage: #e8f0e8;
            --soft-sky: #edf3f6;
        }

        html,
        body,
        [data-testid="stAppViewContainer"] {
            background: var(--soft-bg);
            color: var(--soft-text);
        }

        [data-testid="stAppViewContainer"] > .main {
            background: var(--soft-bg);
        }

        .block-container {
            max-width: 1280px;
            padding-top: 1.6rem;
            padding-bottom: 3rem;
        }

        .app-hero {
            padding: 1.25rem 0 1.1rem;
            border-bottom: 1px solid var(--soft-line);
            margin-bottom: 1.25rem;
        }

        .app-hero h1 {
            margin: 0.15rem 0 0.4rem;
            color: var(--soft-text);
            font-size: 2.25rem;
            line-height: 1.18;
            letter-spacing: 0;
            font-weight: 760;
        }

        .app-hero p {
            margin: 0;
            color: var(--soft-muted);
            font-size: 1.02rem;
        }

        .app-kicker {
            color: var(--soft-accent-dark) !important;
            font-size: 0.86rem !important;
            font-weight: 760;
            letter-spacing: 0 !important;
        }

        [data-testid="stSidebar"] {
            background: #fffaf5;
            border-right: 1px solid var(--soft-line);
        }

        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
            color: var(--soft-text);
        }

        h2, h3 {
            color: var(--soft-text);
            letter-spacing: 0;
        }

        div[data-testid="stTabs"] button {
            color: var(--soft-muted);
            font-weight: 700;
            padding-top: 0.75rem;
            padding-bottom: 0.75rem;
        }

        div[data-testid="stTabs"] button[aria-selected="true"] {
            color: var(--soft-accent-dark);
            border-bottom-color: var(--soft-accent);
        }

        div[data-testid="stExpander"] {
            border: 1px solid var(--soft-line);
            border-radius: 8px;
            background: var(--soft-panel);
        }

        div[data-testid="stFileUploader"] section {
            background: var(--soft-panel-strong);
            border: 1px dashed #d7b9a9;
            border-radius: 8px;
            padding: 1rem;
        }

        div[data-testid="stFileUploader"] section:hover {
            border-color: var(--soft-accent);
            background: #fffaf7;
        }

        div[data-testid="stAlert"] {
            border-radius: 8px;
            border: 1px solid var(--soft-line);
        }

        div[data-testid="stDataFrame"],
        div[data-testid="stTable"] {
            border: 1px solid var(--soft-line);
            border-radius: 8px;
            overflow: hidden;
            background: var(--soft-panel-strong);
        }

        .stButton > button,
        .stDownloadButton > button {
            border-radius: 8px;
            border: 1px solid #c88976;
            background: var(--soft-accent);
            color: #ffffff;
            font-weight: 760;
            box-shadow: none;
            min-height: 2.65rem;
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover {
            border-color: var(--soft-accent-dark);
            background: var(--soft-accent-dark);
            color: #ffffff;
        }

        [data-baseweb="input"] input,
        [data-baseweb="textarea"] textarea,
        [data-baseweb="select"] > div,
        div[data-testid="stNumberInput"] input {
            background-color: var(--soft-panel-strong);
            border-color: #dfcec2;
            color: var(--soft-text);
            border-radius: 8px;
        }

        [data-baseweb="input"] input:focus,
        [data-baseweb="textarea"] textarea:focus,
        [data-baseweb="select"] > div:focus-within,
        div[data-testid="stNumberInput"] input:focus {
            border-color: var(--soft-accent);
            box-shadow: 0 0 0 1px rgba(185, 111, 90, 0.18);
        }

        code,
        pre {
            border-radius: 8px !important;
        }

        .stRadio [role="radiogroup"] {
            background: var(--soft-sky);
            border: 1px solid #dbe6ea;
            border-radius: 8px;
            padding: 0.35rem 0.5rem;
        }

        @media (max-width: 720px) {
            .block-container {
                padding-left: 1rem;
                padding-right: 1rem;
            }

            .app-hero h1 {
                font-size: 1.75rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _load_profile_from_sidebar() -> dict[str, Any]:
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


def _target_template_section(imported_profile: dict[str, Any]) -> dict[str, OutputTemplateConfig]:
    st.subheader("A. OP 最終輸出格式範本")
    st.caption("上傳一份 OP 手打好的最終格式檔；如果裡面有 TINV 和 TPKG 工作表，下面可以分別選。")

    template_file = st.file_uploader(
        "上傳 OP 最終格式範本（Excel / CSV / PDF，可選）",
        type=["xlsx", "xls", "xlsm", "csv", "pdf"],
        key="output_template",
        help="Excel / CSV / PDF",
    )

    candidates: list[TemplateCandidate] = []
    if template_file:
        try:
            candidates = parse_output_template_file(template_file.name, template_file.getvalue())
        except Exception as exc:
            st.error(f"最終格式範本讀取失敗：{exc}")

    configs: dict[str, OutputTemplateConfig] = {}
    if candidates:
        with st.expander("已偵測到的最終範本候選"):
            st.write([candidate.label for candidate in candidates])

        selector_cols = st.columns(2)
        for column, key in zip(selector_cols, DEFAULT_SECTIONS):
            with column:
                selected = _select_template_candidate_for_section(key, candidates)
                if selected:
                    configs[key] = OutputTemplateConfig(
                        key=key,
                        columns=selected.columns,
                        header_row=selected.header_row,
                        data_start_row=selected.data_start_row,
                        label=selected.label,
                    )
                    st.dataframe(
                        build_template_preview(selected.columns, selected.header_row),
                        use_container_width=True,
                        hide_index=True,
                    )

        with st.expander("範本預覽"):
            preview_label = st.selectbox("選擇預覽的偵測結果", [candidate.label for candidate in candidates])
            preview = candidates[[candidate.label for candidate in candidates].index(preview_label)]
            st.dataframe(preview.dataframe.head(50), use_container_width=True)

        return configs

    st.warning("如果最終範本內有 TINV 和 TPKG，建議上傳那份 Excel。沒有範本時可先用手動欄位。")
    manual_cols = st.columns(2)
    for column, key in zip(manual_cols, DEFAULT_SECTIONS):
        with column:
            sheet_profile = _sheet_profile(imported_profile, key)
            default_columns = _columns_from_sheet_profile(sheet_profile) or [
                TemplateColumn(name, index)
                for index, name in enumerate(DEFAULT_TARGET_COLUMNS, start=1)
            ]
            names_text = st.text_area(
                f"{key} 最終輸出欄位（一行一個）",
                value="\n".join(column.name for column in default_columns),
                height=220,
                key=f"{key}_manual_targets",
            )
            names = split_target_columns(names_text)
            configs[key] = OutputTemplateConfig(
                key=key,
                columns=[TemplateColumn(name, index) for index, name in enumerate(names, start=1)],
                header_row=int(sheet_profile.get("output_header_row", 1) or 1),
                data_start_row=int(sheet_profile.get("output_data_start_row", 2) or 2),
                label=f"{key} 手動欄位",
            )

    return configs


def _select_template_candidate_for_section(
    key: str,
    candidates: list[TemplateCandidate],
) -> TemplateCandidate | None:
    labels = ["（不產生）"] + [candidate.label for candidate in candidates]
    default_index = _default_candidate_index(key, candidates)
    selected_label = st.selectbox(
        f"{key} 最終格式",
        labels,
        index=default_index,
        key=f"{key}_template_candidate",
        help=f"選擇 OP 範本中對應 {key} 的工作表或表頭列。",
    )
    if selected_label == "（不產生）":
        return None
    return candidates[[candidate.label for candidate in candidates].index(selected_label)]


def _section_workflow(
    key: str,
    config: OutputTemplateConfig,
    parsed_tables: list[ParsedTable],
    imported_profile: dict[str, Any],
    op_temp_array_mode: bool = False,
) -> tuple[SheetTransformRule | None, dict[str, Any]]:
    defaults = DEFAULT_SECTIONS[key]
    sheet_profile = _sheet_profile(imported_profile, key)
    saved_mappings = _mappings_by_target(sheet_profile.get("mappings", []))

    st.markdown(f"### {defaults['label']}")
    st.caption(f"最終格式來源：{config.label}")

    output_sheet_name = st.text_input(
        "VBA 輸出的工作表名稱",
        value=str(sheet_profile.get("output_sheet", defaults["output_sheet"])),
        key=f"{key}_output_sheet_name",
    )
    fixed_title = st.text_input(
        "固定補入的大標題",
        value=str(sheet_profile.get("fixed_title", defaults["fixed_title"])),
        key=f"{key}_fixed_title",
        help="例如 TINV 需要 HAND TOOL，TPKG 通常可留空。",
    )

    nw_mode = _safe_weight_mode(sheet_profile.get("nw_mode", "source_is_unit"))
    gw_mode = _safe_weight_mode(sheet_profile.get("gw_mode", "source_is_unit"))
    multi_box_mode = bool(sheet_profile.get("multi_box_mode", False))

    if key == "TPKG":
        st.markdown("#### NW / GW 重量推算")
        st.caption("若客戶提供的是整批總重量，系統會先除以箱數後再放進 tempArray，避免 OP 後段報表再乘一次。")
        weight_modes = list(WEIGHT_MODE_LABELS.keys())
        weight_cols = st.columns(2)
        with weight_cols[0]:
            nw_mode = st.selectbox(
                "NW 淨重來源",
                weight_modes,
                index=weight_modes.index(nw_mode),
                format_func=lambda mode: WEIGHT_MODE_LABELS[mode],
                key=f"{key}_nw_mode",
            )
        with weight_cols[1]:
            gw_mode = st.selectbox(
                "GW 毛重來源",
                weight_modes,
                index=weight_modes.index(gw_mode),
                format_func=lambda mode: WEIGHT_MODE_LABELS[mode],
                key=f"{key}_gw_mode",
            )
        multi_box_mode = st.checkbox(
            "啟用多箱處理",
            value=multi_box_mode,
            key=f"{key}_multi_box_mode",
            help="有些客戶需要把每箱與總數量/重量分列顯示；不需要時請關閉。",
        )

    selected_table = _select_source_table_for_section(key, parsed_tables)
    raw_df = selected_table.dataframe
    if selected_table.note:
        st.warning(selected_table.note)
    if raw_df.empty:
        st.error("這個來源沒有資料可預覽。")
        return None, {}

    default_source_sheet_name = str(
        sheet_profile.get("source_sheet")
        or selected_table.source_sheet_name
        or defaults["source_sheet"]
    )
    source_sheet_name = st.text_input(
        "VBA 讀取的來源工作表名稱",
        value=default_source_sheet_name,
        key=f"{key}_source_sheet_name_{_safe_file_stem(default_source_sheet_name)[:40]}",
        disabled=selected_table.kind != "excel",
        help="會自動帶入目前選到的 Excel 來源分頁；貼回轉檔工具時，原始資料分頁名稱要和這裡一致。",
    )

    st.dataframe(raw_df.head(45), use_container_width=True)

    max_rows = max(len(raw_df), 1)
    auto_header_row = suggest_header_row(raw_df)
    auto_data_start_row = min(max_rows, auto_header_row + 1)
    st.caption(f"系統預估表頭在第 {auto_header_row} 列，資料從第 {auto_data_start_row} 列開始；如果預覽不對，再手動調整。")
    setup_cols = st.columns(4)
    with setup_cols[0]:
        header_row = st.number_input(
            "來源表頭列",
            min_value=1,
            max_value=max_rows,
            value=_bounded_int(sheet_profile.get("header_row", auto_header_row), 1, max_rows),
            step=1,
            key=f"{key}_header_row",
        )
    with setup_cols[1]:
        data_start_row = st.number_input(
            "來源資料開始列",
            min_value=1,
            max_value=max_rows,
            value=_bounded_int(sheet_profile.get("data_start_row", auto_data_start_row), 1, max_rows),
            step=1,
            key=f"{key}_data_start_row",
        )
    with setup_cols[2]:
        output_header_row = st.number_input(
            "最終格式表頭列",
            min_value=1,
            max_value=500,
            value=_bounded_int(sheet_profile.get("output_header_row", config.header_row), 1, 500),
            step=1,
            key=f"{key}_output_header_row",
        )
    with setup_cols[3]:
        output_data_start_row = st.number_input(
            "最終格式資料開始列",
            min_value=1,
            max_value=500,
            value=_bounded_int(sheet_profile.get("output_data_start_row", config.data_start_row), 1, 500),
            step=1,
            key=f"{key}_output_data_start_row",
        )

    structured_df, source_headers = prepare_structured_table(raw_df, int(header_row), int(data_start_row))
    source_options = build_source_options(source_headers)
    with st.expander("用來源表頭列整理後的預覽"):
        st.dataframe(structured_df.head(30), use_container_width=True)

    st.markdown("欄位對應")
    target_columns = _target_columns_for_section(key, config, op_temp_array_mode)
    if op_temp_array_mode:
        st.info("OP tempArray 模式會使用內部 TINV / TPKG 標準欄位做對應，產出的 VBA 仍會寫回 Tinv / Tpkg 固定格式。")
    mappings = _mapping_editor(
        section_key=key,
        target_columns=target_columns,
        source_options=source_options,
        saved_mappings=saved_mappings,
    )

    rule = SheetTransformRule(
        procedure_name=str(defaults["procedure"]),
        source_sheet_name=source_sheet_name.strip() or str(defaults["source_sheet"]),
        output_sheet_name=output_sheet_name.strip() or str(defaults["output_sheet"]),
        mappings=mappings,
        header_row=int(header_row),
        data_start_row=int(data_start_row),
        output_header_row=int(output_header_row),
        output_data_start_row=int(output_data_start_row),
        fixed_title=fixed_title.strip(),
        nw_mode=nw_mode,
        gw_mode=gw_mode,
        multi_box_mode=multi_box_mode,
    )
    profile = {
        "source_sheet": rule.source_sheet_name,
        "output_sheet": rule.output_sheet_name,
        "header_row": rule.header_row,
        "data_start_row": rule.data_start_row,
        "output_header_row": rule.output_header_row,
        "output_data_start_row": rule.output_data_start_row,
        "fixed_title": rule.fixed_title,
        "nw_mode": rule.nw_mode,
        "gw_mode": rule.gw_mode,
        "multi_box_mode": rule.multi_box_mode,
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
    return rule, profile


def _op_temp_array_runtime_rule(key: str, rule: SheetTransformRule) -> SheetTransformRule:
    runtime = OP_TEMP_ARRAY_RUNTIME_RULES.get(key)
    if not runtime:
        return rule

    mappings = [
        FieldMapping(target=target, source_header="", source_index=source_index)
        for target, source_index in runtime["mappings"]
    ]

    return SheetTransformRule(
        procedure_name=rule.procedure_name,
        source_sheet_name=str(DEFAULT_SECTIONS[key]["source_sheet"]),
        output_sheet_name=rule.output_sheet_name,
        mappings=mappings,
        header_row=int(runtime["header_row"]),
        data_start_row=int(runtime["data_start_row"]),
        output_header_row=rule.output_header_row,
        output_data_start_row=rule.output_data_start_row,
        fixed_title=rule.fixed_title,
        nw_mode=rule.nw_mode,
        gw_mode=rule.gw_mode,
        multi_box_mode=rule.multi_box_mode,
    )


def _profile_from_rule(rule: SheetTransformRule) -> dict[str, Any]:
    return {
        "source_sheet": rule.source_sheet_name,
        "output_sheet": rule.output_sheet_name,
        "header_row": rule.header_row,
        "data_start_row": rule.data_start_row,
        "output_header_row": rule.output_header_row,
        "output_data_start_row": rule.output_data_start_row,
        "fixed_title": rule.fixed_title,
        "nw_mode": rule.nw_mode,
        "gw_mode": rule.gw_mode,
        "multi_box_mode": rule.multi_box_mode,
        "mappings": [
            {
                "target": mapping.target,
                "target_col": mapping.target_col,
                "source_header": mapping.source_header,
                "source_index": mapping.source_index,
            }
            for mapping in rule.mappings
        ],
    }


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


def _select_source_table_for_section(key: str, parsed_tables: list[ParsedTable]) -> ParsedTable:
    labels = [table.label for table in parsed_tables]
    default_index = _default_source_table_index(key, parsed_tables)
    selected_label = st.selectbox(
        f"選擇 {key} 的客戶原始來源",
        labels,
        index=default_index,
        key=f"{key}_source_table",
    )
    return parsed_tables[labels.index(selected_label)]


def _mapping_editor(
    section_key: str,
    target_columns: list[TemplateColumn],
    source_options: list[SourceColumn],
    saved_mappings: dict[str, dict[str, Any]],
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
                    key=f"{section_key}_map_{target_index}_{target.name}_{target.column_index}",
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


def _default_candidate_index(key: str, candidates: list[TemplateCandidate]) -> int:
    preferred = {
        "TINV": ["tinv", "invoice", "inv"],
        "TPKG": ["tpkg", "packing", "pack", "pkg"],
    }.get(key, [key.lower()])

    def score_candidate(candidate: TemplateCandidate) -> int:
        sheet = candidate.sheet_name.lower()
        label = candidate.label.lower()
        columns = " ".join(column.name.lower() for column in candidate.columns)
        score = 0

        for rank, needle in enumerate(preferred):
            weight = len(preferred) - rank
            if needle in sheet:
                score += 100 * weight
            if needle in label:
                score += 40 * weight
            if needle in columns:
                score += 12 * weight

        if key == "TPKG":
            if any(word in columns for word in ["packing no", "ctn", "carton", "gross weight", "net weight", "measurement"]):
                score += 80
            if "invoice" in sheet and "packing" not in sheet:
                score -= 200

        if key == "TINV":
            if "unit price" in columns and "amount" in columns:
                score += 100
            elif any(word in columns for word in ["unit price", "amount", "marks"]):
                score += 40
            if any(word in columns for word in ["packing no", "net weight", "gross weight", "measurement"]):
                score -= 120
            if "packing" in sheet and "invoice" not in sheet:
                score -= 200

        score += candidate.score
        return score

    best_index = 0
    best_score = -10**9
    for index, candidate in enumerate(candidates, start=1):
        score = score_candidate(candidate)
        if score > best_score:
            best_index = index
            best_score = score

    return best_index if candidates else 0


def _default_source_table_index(key: str, parsed_tables: list[ParsedTable]) -> int:
    preferred = (
        ["invoice", "inv", "unit price", "amount"]
        if key == "TINV"
        else ["packing", "pack", "pkg", "roll/no", "net weight", "gross weight", "ctn"]
    )
    penalties = (
        ["packing", "net weight", "gross weight"]
        if key == "TINV"
        else ["invoice", "unit price", "amount"]
    )

    best_index = 0
    best_score = -10**9
    for index, table in enumerate(parsed_tables):
        preview_values = []
        if not table.dataframe.empty:
            preview_values = table.dataframe.head(25).astype(str).values.ravel().tolist()
        haystack = " ".join([table.label, *preview_values]).lower()
        sheet_key = (
            table.source_sheet_name.lower()
            .replace("+", "")
            .replace(" ", "")
            .replace("_", "")
            .replace("-", "")
        )
        score = 0
        if sheet_key in {"invpkg", "invandpkg", "invpacking", "invoicepacking"}:
            score += 500
        if sheet_key in {"tinv", "tpkg", "menu", "hs"}:
            score -= 500
        for rank, needle in enumerate(preferred):
            if needle in haystack:
                score += (len(preferred) - rank) * 20
        for needle in penalties:
            if needle in haystack:
                score -= 30
        if score > best_score:
            best_index = index
            best_score = score

    return best_index


def _default_source_option_index(
    target: str,
    source_options: list[SourceColumn],
    saved_mapping: dict[str, Any],
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


def _sheet_profile(imported_profile: dict[str, Any], key: str) -> dict[str, Any]:
    sheets = imported_profile.get("sheets", {})
    if isinstance(sheets, dict) and isinstance(sheets.get(key), dict):
        return sheets[key]
    return {}


def _classification_profile(imported_profile: dict[str, Any]) -> dict[str, Any]:
    classification = imported_profile.get("classification", {})
    if not isinstance(classification, dict):
        classification = {}

    tags = classification.get("tags", [])
    if isinstance(tags, str):
        tags = _split_tags(tags)
    elif not isinstance(tags, list):
        tags = []

    return {
        "end_customer_name": str(classification.get("end_customer_name", "")),
        "rule_category": str(classification.get("rule_category", "")),
        "tags": [str(tag) for tag in tags if str(tag).strip()],
        "note": str(classification.get("note", "")),
    }


def _columns_from_sheet_profile(sheet_profile: dict[str, Any]) -> list[TemplateColumn]:
    mappings = sheet_profile.get("mappings", [])
    if not isinstance(mappings, list):
        return []

    columns: list[TemplateColumn] = []
    for index, mapping in enumerate(mappings, start=1):
        if not isinstance(mapping, dict):
            continue
        target = str(mapping.get("target", "")).strip()
        if not target:
            continue
        target_col = int(mapping.get("target_col") or index)
        columns.append(TemplateColumn(target, target_col))
    return columns


def _mappings_by_target(mappings: Any) -> dict[str, dict[str, Any]]:
    if not isinstance(mappings, list):
        return {}
    result: dict[str, dict[str, Any]] = {}
    for mapping in mappings:
        if isinstance(mapping, dict) and mapping.get("target"):
            result[str(mapping["target"])] = mapping
    return result


def _show_next_steps() -> None:
    with st.expander("建議工作流程"):
        st.markdown(
            "1. 上傳 OP 以前手打好的最終輸出格式，裡面可同時有 TINV 和 TPKG。\n"
            "2. 分別選 TINV 最終格式與 TPKG 最終格式。\n"
            "3. 上傳客戶原始 invoice / packing list。\n"
            "4. 在 TINV 分頁選 invoice 來源，在 TPKG 分頁選 packing list 來源。\n"
            "5. 分別設定來源表頭列、來源資料開始列與欄位對應。\n"
            "6. 下載同一份 VBA 和客戶規則 JSON。"
        )


def _bounded_int(value: object, minimum: int, maximum: int) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        number = minimum
    return max(minimum, min(maximum, number))


def _safe_weight_mode(value: object) -> str:
    mode = str(value or "source_is_unit")
    if mode in WEIGHT_MODE_LABELS:
        return mode
    return "source_is_unit"


def _split_tags(value: object) -> list[str]:
    if isinstance(value, list):
        raw_items = value
    else:
        raw_items = str(value or "").replace("，", ",").replace("、", ",").split(",")

    tags: list[str] = []
    for item in raw_items:
        tag = str(item).strip()
        if tag and tag not in tags:
            tags.append(tag)
    return tags


def _safe_file_stem(value: str) -> str:
    cleaned = "".join(char if char.isalnum() or char in "-_" else "_" for char in value.strip())
    return cleaned.strip("_") or "customer_rule"


if __name__ == "__main__":
    main()
