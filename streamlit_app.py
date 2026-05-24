from __future__ import annotations

from html import escape
import tempfile
from pathlib import Path
from typing import Any

import streamlit as st

from generate_itinerary import (
    DEFAULT_CONFIG,
    DEFAULT_TEMPLATE,
    DEFAULT_TRIP_FIELDS,
    fill_document,
    format_chinese_date,
    load_config,
    output_path_from_config,
    read_order_data,
)


CSS = """
<style>
:root {
  --ink: #1d2422;
  --muted: #63706b;
  --line: #d7ddd7;
  --paper: #fbfaf6;
  --panel: #ffffff;
  --accent: #0f7b68;
  --accent-2: #c86f3d;
  --focus: #188f79;
}

.stApp {
  background:
    linear-gradient(90deg, rgba(15, 123, 104, 0.04) 1px, transparent 1px),
    linear-gradient(rgba(15, 123, 104, 0.035) 1px, transparent 1px),
    var(--paper);
  background-size: 28px 28px;
  color: var(--ink);
}

.block-container {
  max-width: 880px;
  padding-top: 42px;
  padding-bottom: 56px;
}

[data-testid="stHeader"] {
  background: transparent;
}

.tool-head {
  border-left: 8px solid var(--accent);
  padding: 6px 0 8px 18px;
  margin-bottom: 26px;
}

.tool-head h1 {
  font-size: 2.4rem;
  line-height: 1.08;
  margin: 0 0 8px 0;
  letter-spacing: 0;
  color: var(--ink);
}

.tool-head p {
  margin: 0;
  color: var(--muted);
  font-size: 1rem;
}

.step-label {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  margin: 18px 0 8px;
  color: var(--accent);
  font-weight: 700;
}

.field-note {
  color: var(--muted);
  font-size: 0.9rem;
  margin: -2px 0 8px;
}

[data-testid="stWidgetLabel"] p {
  color: var(--ink);
  font-weight: 700;
  font-size: 0.95rem;
}

.stTextInput input,
[data-testid="stTextInput"] input,
[data-baseweb="input"] input {
  background: var(--panel) !important;
  color: var(--ink) !important;
  border: 1px solid var(--line) !important;
  border-radius: 6px;
  min-height: 44px;
  box-shadow: none !important;
}

.stTextInput input:focus,
[data-testid="stTextInput"] input:focus,
[data-baseweb="input"] input:focus {
  border-color: var(--focus) !important;
  box-shadow: 0 0 0 3px rgba(15, 123, 104, 0.14) !important;
}

.stTextInput input::placeholder {
  color: #8a9691 !important;
}

.stButton > button,
.stDownloadButton > button {
  border-radius: 6px;
  min-height: 44px;
  font-weight: 700;
}

.stDownloadButton > button {
  background: var(--accent);
  color: white;
  border-color: var(--accent);
}

[data-testid="stFileUploader"] {
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 8px;
  padding: 8px 14px 14px;
}

[data-testid="stFileUploader"] section {
  background: #f6f8f5 !important;
  border: 1px dashed #9eb1aa !important;
  border-radius: 6px !important;
}

[data-testid="stFileUploader"] button {
  background: var(--accent) !important;
  color: white !important;
  border-color: var(--accent) !important;
  border-radius: 6px !important;
}

[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] span,
[data-testid="stFileUploader"] p {
  color: var(--ink) !important;
}

[data-testid="stAlert"] {
  background: #e8f4ff;
  color: #19425f;
  border-radius: 6px;
}

[data-testid="stMetricValue"] {
  color: var(--accent-2);
}

.preview-wrap {
  margin-top: 18px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.84);
  overflow: hidden;
}

.preview-head {
  display: flex;
  justify-content: space-between;
  gap: 16px;
  padding: 16px 18px;
  border-bottom: 1px solid var(--line);
  background: #f6f8f5;
}

.preview-title {
  font-size: 1.18rem;
  font-weight: 800;
  color: var(--ink);
  margin: 0 0 6px;
}

.preview-file {
  color: var(--muted);
  font-size: 0.88rem;
  word-break: break-all;
}

.preview-count {
  min-width: 76px;
  text-align: right;
}

.preview-count b {
  display: block;
  color: var(--accent-2);
  font-size: 1.8rem;
  line-height: 1;
}

.preview-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 0;
}

.preview-item {
  padding: 12px 18px;
  border-bottom: 1px solid var(--line);
}

.preview-item:nth-child(odd) {
  border-right: 1px solid var(--line);
}

.preview-item.wide {
  grid-column: 1 / -1;
  border-right: 0;
}

.preview-label {
  color: var(--muted);
  font-size: 0.82rem;
  margin-bottom: 4px;
}

.preview-value {
  color: var(--ink);
  font-weight: 650;
  line-height: 1.55;
  word-break: break-word;
}

@media (max-width: 720px) {
  .preview-head,
  .preview-grid {
    display: block;
  }

  .preview-count {
    margin-top: 12px;
    text-align: left;
  }

  .preview-item,
  .preview-item:nth-child(odd) {
    border-right: 0;
  }
}
</style>
"""


def app_config_defaults() -> dict[str, Any]:
    if DEFAULT_CONFIG.exists():
        try:
            return load_config(DEFAULT_CONFIG)
        except SystemExit:
            pass
    return {
        **DEFAULT_TRIP_FIELDS,
        "route": "",
        "leaders": "",
        "leader_phone": "",
    }


def write_uploaded_file(uploaded_file: Any, destination: Path) -> None:
    destination.write_bytes(uploaded_file.getbuffer())


def prepare_itinerary(
    uploaded_excel: Any,
    route: str,
    leaders: str,
    leader_phone: str,
    operator: str,
    operator_phone: str,
) -> tuple[dict[str, Any], list[dict[str, str]], str]:
    with tempfile.TemporaryDirectory() as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        excel_path = temp_dir / "order.xlsx"
        write_uploaded_file(uploaded_excel, excel_path)

        order_data = read_order_data(excel_path)
        config = {
            **DEFAULT_TRIP_FIELDS,
            "route": route.strip(),
            "leaders": leaders.strip(),
            "leader_phone": leader_phone.strip(),
            "operator": operator.strip(),
            "operator_phone": operator_phone.strip(),
        }
        if order_data["start_date"]:
            config["start_date"] = order_data["start_date"]
        if order_data["end_date"]:
            config["end_date"] = order_data["end_date"]
        if order_data["route_name"]:
            config["route_name"] = order_data["route_name"]

        output_name = output_path_from_config(config, None).name
        return config, order_data["passengers"], output_name


def build_docx_from_data(
    config: dict[str, Any],
    passengers: list[dict[str, str]],
    output_name: str,
    template_path: Path,
) -> bytes:
    with tempfile.TemporaryDirectory() as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        output_path = temp_dir / output_name
        fill_document(template_path, output_path, config, passengers)
        return output_path.read_bytes()


def build_docx(
    uploaded_excel: Any,
    route: str,
    leaders: str,
    leader_phone: str,
    operator: str,
    operator_phone: str,
    template_path: Path,
) -> tuple[bytes, str, int, dict[str, Any], list[dict[str, str]]]:
    config, passengers, output_name = prepare_itinerary(
        uploaded_excel,
        route,
        leaders,
        leader_phone,
        operator,
        operator_phone,
    )
    docx_bytes = build_docx_from_data(config, passengers, output_name, template_path)
    return docx_bytes, output_name, len(passengers), config, passengers


def render_preview(config: dict[str, Any], passengers: list[dict[str, str]], filename: str) -> None:
    start_date = format_chinese_date(str(config.get("start_date") or ""))
    end_date = format_chinese_date(str(config.get("end_date") or ""))
    travel_period = f"{start_date} 至 {end_date}" if start_date and end_date else "未识别"
    route = escape(str(config.get("route") or ""))
    leaders = escape(str(config.get("leaders") or ""))
    leader_phone = escape(str(config.get("leader_phone") or ""))
    operator = escape(str(config.get("operator") or ""))
    operator_phone = escape(str(config.get("operator_phone") or ""))
    st.markdown(
        f"""
        <div class="preview-wrap">
          <div class="preview-head">
            <div>
              <p class="preview-title">行程单预览</p>
              <div class="preview-file">{escape(filename)}</div>
            </div>
            <div class="preview-count">
              <span>旅客人数</span>
              <b>{len(passengers)}</b>
            </div>
          </div>
          <div class="preview-grid">
            <div class="preview-item wide">
              <div class="preview-label">旅游路线</div>
              <div class="preview-value">{route}</div>
            </div>
            <div class="preview-item">
              <div class="preview-label">出行时间</div>
              <div class="preview-value">{escape(travel_period)}</div>
            </div>
            <div class="preview-item">
              <div class="preview-label">领队</div>
              <div class="preview-value">{leaders}</div>
            </div>
            <div class="preview-item">
              <div class="preview-label">领队电话</div>
              <div class="preview-value">{leader_phone}</div>
            </div>
            <div class="preview-item">
              <div class="preview-label">计调</div>
              <div class="preview-value">{operator} / {operator_phone}</div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.dataframe(passengers, use_container_width=True, hide_index=True)


def main() -> None:
    st.set_page_config(
        page_title="行程单生成工具",
        layout="centered",
        initial_sidebar_state="collapsed",
    )

    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(
        """
        <div class="tool-head">
          <h1>行程单生成工具</h1>
          <p>上传订单 Excel，填写领队信息，生成可编辑 Word 行程单。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    defaults = app_config_defaults()

    st.markdown('<div class="step-label">01 上传订单</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="field-note">上传从订单系统导出的 Excel 文件，文件里需要有“订单数据”工作表。</div>',
        unsafe_allow_html=True,
    )
    uploaded_excel = st.file_uploader(
        "订单 Excel",
        type=["xlsx"],
        label_visibility="visible",
        help="表格中需要包含工作表“订单数据”，以及旅客姓名、旅客证件类型、旅客证件号码、旅客手机号码列。",
    )

    st.markdown('<div class="step-label">02 填写行程信息</div>', unsafe_allow_html=True)
    route = st.text_input(
        "旅游路线",
        value=str(defaults.get("route") or ""),
        placeholder="例如：西宁-祁连-七彩丹霞-敦煌-青海湖-西宁",
    )
    col1, col2 = st.columns(2)
    with col1:
        leaders = st.text_input(
            "领队",
            value=str(defaults.get("leaders") or ""),
            placeholder="多个领队用顿号分隔",
        )
    with col2:
        leader_phone = st.text_input(
            "领队电话",
            value=str(defaults.get("leader_phone") or ""),
            placeholder="填写手机号",
        )
    col3, col4 = st.columns(2)
    with col3:
        operator = st.text_input(
            "计调",
            value=str(defaults.get("operator") or DEFAULT_TRIP_FIELDS["operator"]),
            placeholder="填写计调姓名",
        )
    with col4:
        operator_phone = st.text_input(
            "计调电话",
            value=str(defaults.get("operator_phone") or DEFAULT_TRIP_FIELDS["operator_phone"]),
            placeholder="填写计调电话",
        )

    st.markdown('<div class="step-label">03 生成下载</div>', unsafe_allow_html=True)
    can_generate = bool(
        uploaded_excel
        and route.strip()
        and leaders.strip()
        and leader_phone.strip()
        and operator.strip()
        and operator_phone.strip()
    )
    generate = st.button("生成行程单", type="primary", disabled=not can_generate, use_container_width=True)

    if not DEFAULT_TEMPLATE.exists():
        st.error(f"找不到模板文件：{DEFAULT_TEMPLATE}")
    elif generate:
        try:
            docx_bytes, filename, passenger_count, preview_config, preview_passengers = build_docx(
                uploaded_excel,
                route,
                leaders,
                leader_phone,
                operator,
                operator_phone,
                DEFAULT_TEMPLATE,
            )
        except SystemExit as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"生成失败：{exc}")
        else:
            st.success("已生成行程单")
            metric_col, docx_col = st.columns([1, 3])
            with metric_col:
                st.metric("旅客人数", passenger_count)
            with docx_col:
                st.download_button(
                    "下载 Word 行程单",
                    data=docx_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            render_preview(preview_config, preview_passengers, filename)
    elif not can_generate:
        st.info("上传 Excel 并填写行程字段后即可生成。")


if __name__ == "__main__":
    main()
