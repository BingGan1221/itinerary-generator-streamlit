from __future__ import annotations

from io import BytesIO
import tempfile
from pathlib import Path
from typing import Any

import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from generate_itinerary import (
    DEFAULT_CONFIG,
    DEFAULT_TEMPLATE,
    DEFAULT_TRIP_FIELDS,
    format_chinese_date,
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
    template_path: Path,
) -> tuple[bytes, str, int]:
    config, passengers, output_name = prepare_itinerary(uploaded_excel, route, leaders, leader_phone)
    docx_bytes = build_docx_from_data(config, passengers, output_name, template_path)
    return docx_bytes, output_name, len(passengers)


def build_pdf(
    config: dict[str, Any],
    passengers: list[dict[str, str]],
    output_name: str,
) -> tuple[bytes, str]:
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ChineseTitle",
        parent=styles["Title"],
        fontName="STSong-Light",
        fontSize=18,
        leading=24,
        alignment=TA_CENTER,
        spaceAfter=10,
    )
    normal_style = ParagraphStyle(
        "ChineseNormal",
        parent=styles["Normal"],
        fontName="STSong-Light",
        fontSize=9,
        leading=13,
    )

    buffer = BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=14 * mm,
        leftMargin=14 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
    )

    travel_dates = " 至 ".join(
        value
        for value in (
            format_chinese_date(str(config.get("start_date") or "")),
            format_chinese_date(str(config.get("end_date") or "")),
        )
        if value
    )
    summary_data = [
        ["旅游路线", Paragraph(str(config.get("route") or ""), normal_style)],
        ["旅行社", str(config.get("agency") or "")],
        ["客源地", str(config.get("source") or "")],
        ["出行时间", travel_dates],
        ["领队", str(config.get("leaders") or "")],
        ["领队电话", str(config.get("leader_phone") or "")],
        ["操作人", str(config.get("operator") or "")],
        ["操作电话", str(config.get("operator_phone") or "")],
        ["旅客人数", str(len(passengers))],
    ]
    summary_table = Table(summary_data, colWidths=[28 * mm, 140 * mm])
    summary_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "STSong-Light"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#eef5f2")),
                ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#0f7b68")),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd7d2")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )

    passenger_rows = [["姓名", "证件类型", "证件号码", "手机号码"]]
    passenger_rows.extend(
        [
            passenger["旅客姓名"],
            passenger["旅客证件类型"],
            passenger["旅客证件号码"],
            passenger["旅客手机号码"],
        ]
        for passenger in passengers
    )
    passenger_table = Table(passenger_rows, repeatRows=1, colWidths=[25 * mm, 30 * mm, 70 * mm, 36 * mm])
    passenger_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "STSong-Light"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f7b68")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd7d2")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )

    story = [
        Paragraph("旅游团队行程单", title_style),
        summary_table,
        Spacer(1, 8 * mm),
        Paragraph("旅客名单", normal_style),
        Spacer(1, 3 * mm),
        passenger_table,
    ]
    document.build(story)
    return buffer.getvalue(), Path(output_name).with_suffix(".pdf").name


def build_files(
    uploaded_excel: Any,
    route: str,
    leaders: str,
    leader_phone: str,
    template_path: Path,
) -> tuple[bytes, str, bytes, str, int]:
    config, passengers, output_name = prepare_itinerary(uploaded_excel, route, leaders, leader_phone)
    docx_bytes = build_docx_from_data(config, passengers, output_name, template_path)
    pdf_bytes, pdf_name = build_pdf(config, passengers, output_name)
    return docx_bytes, output_name, pdf_bytes, pdf_name, len(passengers)


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
          <p>上传订单 Excel，填写领队信息，生成可编辑 Word 和 PDF 行程单。</p>
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

    st.markdown('<div class="step-label">03 生成下载</div>', unsafe_allow_html=True)
    can_generate = bool(uploaded_excel and route.strip() and leaders.strip() and leader_phone.strip())
    generate = st.button("生成行程单", type="primary", disabled=not can_generate, use_container_width=True)

    if not DEFAULT_TEMPLATE.exists():
        st.error(f"找不到模板文件：{DEFAULT_TEMPLATE}")
    elif generate:
        try:
            docx_bytes, filename, pdf_bytes, pdf_filename, passenger_count = build_files(
                uploaded_excel, route, leaders, leader_phone, DEFAULT_TEMPLATE
            )
        except SystemExit as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"生成失败：{exc}")
        else:
            st.success("已生成行程单")
            metric_col, docx_col, pdf_col = st.columns([1, 2, 2])
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
            with pdf_col:
                st.download_button(
                    "下载 PDF 行程单",
                    data=pdf_bytes,
                    file_name=pdf_filename,
                    mime="application/pdf",
                    use_container_width=True,
                )
    elif not can_generate:
        st.info("上传 Excel 并填写三个行程字段后即可生成。")


if __name__ == "__main__":
    main()
