from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any

import streamlit as st

from generate_itinerary import (
    DEFAULT_CONFIG,
    DEFAULT_TEMPLATE,
    DEFAULT_TRIP_FIELDS,
    fill_document,
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

.stTextInput input {
  border-radius: 6px;
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
  background: rgba(255, 255, 255, 0.72);
  border: 1px solid var(--line);
  border-radius: 8px;
  padding: 4px 14px 14px;
}

[data-testid="stMetricValue"] {
  color: var(--accent-2);
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


def build_docx(
    uploaded_excel: Any,
    route: str,
    leaders: str,
    leader_phone: str,
    template_path: Path,
) -> tuple[bytes, str, int]:
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
        }
        if order_data["start_date"]:
            config["start_date"] = order_data["start_date"]
        if order_data["end_date"]:
            config["end_date"] = order_data["end_date"]
        if order_data["route_name"]:
            config["route_name"] = order_data["route_name"]

        output_name = output_path_from_config(config, None).name
        output_path = temp_dir / output_name
        fill_document(template_path, output_path, config, order_data["passengers"])
        return output_path.read_bytes(), output_name, len(order_data["passengers"])


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
    uploaded_excel = st.file_uploader(
        "订单 Excel",
        type=["xlsx"],
        label_visibility="collapsed",
        help="表格中需要包含工作表“订单数据”，以及旅客姓名、旅客证件类型、旅客证件号码、旅客手机号码列。",
    )

    st.markdown('<div class="step-label">02 填写行程信息</div>', unsafe_allow_html=True)
    route = st.text_input("旅游路线", value=str(defaults.get("route") or ""))
    col1, col2 = st.columns(2)
    with col1:
        leaders = st.text_input("领队", value=str(defaults.get("leaders") or ""))
    with col2:
        leader_phone = st.text_input("领队电话", value=str(defaults.get("leader_phone") or ""))

    st.markdown('<div class="step-label">03 生成下载</div>', unsafe_allow_html=True)
    can_generate = bool(uploaded_excel and route.strip() and leaders.strip() and leader_phone.strip())
    generate = st.button("生成行程单", type="primary", disabled=not can_generate, use_container_width=True)

    if not DEFAULT_TEMPLATE.exists():
        st.error(f"找不到模板文件：{DEFAULT_TEMPLATE}")
    elif generate:
        try:
            docx_bytes, filename, passenger_count = build_docx(
                uploaded_excel,
                route,
                leaders,
                leader_phone,
                DEFAULT_TEMPLATE,
            )
        except SystemExit as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"生成失败：{exc}")
        else:
            st.success("已生成行程单")
            metric_col, download_col = st.columns([1, 2])
            with metric_col:
                st.metric("旅客人数", passenger_count)
            with download_col:
                st.download_button(
                    "下载 Word 行程单",
                    data=docx_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
    elif not can_generate:
        st.info("上传 Excel 并填写三个行程字段后即可生成。")


if __name__ == "__main__":
    main()
