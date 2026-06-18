from __future__ import annotations

from html import escape
import json
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


ROUTE_CLIPBOARD = DEFAULT_CONFIG.parent / "route_clipboard.json"
LEADER_CLIPBOARD = DEFAULT_CONFIG.parent / "leader_clipboard.json"

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


def normalize_route_clipboard(routes: Any) -> list[str]:
    if not isinstance(routes, list):
        return []

    normalized: list[str] = []
    seen: set[str] = set()
    for route in routes:
        route_text = str(route or "").strip()
        if route_text and route_text not in seen:
            normalized.append(route_text)
            seen.add(route_text)
    return normalized


def load_route_clipboard() -> list[str]:
    if not ROUTE_CLIPBOARD.exists():
        return []
    try:
        with ROUTE_CLIPBOARD.open("r", encoding="utf-8") as f:
            routes = json.load(f)
    except json.JSONDecodeError:
        st.warning(f"路线库文件不是合法 JSON，已暂时忽略：{ROUTE_CLIPBOARD}")
        return []
    except OSError as exc:
        st.warning(f"读取路线库失败，已暂时忽略：{exc}")
        return []

    if not isinstance(routes, list):
        st.warning(f"路线库需要是 JSON 数组，已暂时忽略：{ROUTE_CLIPBOARD}")
        return []
    return normalize_route_clipboard(routes)


def save_route_clipboard(routes: list[str]) -> None:
    ROUTE_CLIPBOARD.parent.mkdir(parents=True, exist_ok=True)
    with ROUTE_CLIPBOARD.open("w", encoding="utf-8") as f:
        json.dump(normalize_route_clipboard(routes), f, ensure_ascii=False, indent=2)
        f.write("\n")


def normalize_leader_clipboard(leaders: Any) -> list[dict[str, str]]:
    if not isinstance(leaders, list):
        return []

    normalized: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for leader in leaders:
        if not isinstance(leader, dict):
            continue
        leader_names = str(leader.get("leaders") or "").strip()
        leader_phone = str(leader.get("leader_phone") or "").strip()
        if not leader_names or not leader_phone:
            continue
        key = (leader_names, leader_phone)
        if key in seen:
            continue
        normalized.append({"leaders": leader_names, "leader_phone": leader_phone})
        seen.add(key)
    return normalized


def load_leader_clipboard() -> list[dict[str, str]]:
    if not LEADER_CLIPBOARD.exists():
        return []
    try:
        with LEADER_CLIPBOARD.open("r", encoding="utf-8") as f:
            leaders = json.load(f)
    except json.JSONDecodeError:
        st.warning(f"领队库文件不是合法 JSON，已暂时忽略：{LEADER_CLIPBOARD}")
        return []
    except OSError as exc:
        st.warning(f"读取领队库失败，已暂时忽略：{exc}")
        return []

    if not isinstance(leaders, list):
        st.warning(f"领队库需要是 JSON 数组，已暂时忽略：{LEADER_CLIPBOARD}")
        return []
    return normalize_leader_clipboard(leaders)


def save_leader_clipboard(leaders: list[dict[str, str]]) -> None:
    LEADER_CLIPBOARD.parent.mkdir(parents=True, exist_ok=True)
    with LEADER_CLIPBOARD.open("w", encoding="utf-8") as f:
        json.dump(normalize_leader_clipboard(leaders), f, ensure_ascii=False, indent=2)
        f.write("\n")


def leader_clipboard_label(leader: dict[str, str]) -> str:
    return f"{leader['leaders']} / {leader['leader_phone']}"


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
    if "route" not in st.session_state:
        st.session_state["route"] = str(defaults.get("route") or "")
    if "leaders" not in st.session_state:
        st.session_state["leaders"] = str(defaults.get("leaders") or "")
    if "leader_phone" not in st.session_state:
        st.session_state["leader_phone"] = str(defaults.get("leader_phone") or "")
    if "route_clipboard" not in st.session_state:
        st.session_state["route_clipboard"] = load_route_clipboard()
    if "leader_clipboard" not in st.session_state:
        st.session_state["leader_clipboard"] = load_leader_clipboard()

    def use_selected_route() -> None:
        selected_route = str(st.session_state.get("route_clipboard_selected") or "").strip()
        if selected_route:
            st.session_state["route"] = selected_route
            st.session_state["route_clipboard_message"] = ("success", "已填入选中的旅游路线。")

    def save_current_route() -> None:
        current_route = str(st.session_state.get("route") or "").strip()
        if not current_route:
            st.session_state["route_clipboard_message"] = ("warning", "旅游路线为空，无法保存。")
            return

        routes = normalize_route_clipboard(st.session_state.get("route_clipboard", []))
        if current_route in routes:
            st.session_state["route_clipboard_selected"] = current_route
            st.session_state["route_clipboard_message"] = ("info", "这条旅游路线已经在路线库里。")
            return

        routes.append(current_route)
        save_route_clipboard(routes)
        st.session_state["route_clipboard"] = routes
        st.session_state["route_clipboard_selected"] = current_route
        st.session_state["route_clipboard_message"] = ("success", "已保存当前旅游路线。")

    def delete_selected_route() -> None:
        selected_route = str(st.session_state.get("route_clipboard_selected") or "").strip()
        if not selected_route:
            st.session_state["route_clipboard_message"] = ("warning", "请先选择要删除的旅游路线。")
            return

        routes = [
            route
            for route in normalize_route_clipboard(st.session_state.get("route_clipboard", []))
            if route != selected_route
        ]
        save_route_clipboard(routes)
        st.session_state["route_clipboard"] = routes
        if routes:
            st.session_state["route_clipboard_selected"] = routes[0]
        else:
            st.session_state.pop("route_clipboard_selected", None)
        st.session_state["route_clipboard_message"] = ("success", "已删除选中的旅游路线。")

    def use_selected_leader() -> None:
        selected_label = str(st.session_state.get("leader_clipboard_selected") or "").strip()
        leaders = normalize_leader_clipboard(st.session_state.get("leader_clipboard", []))
        selected_leader = next(
            (leader for leader in leaders if leader_clipboard_label(leader) == selected_label),
            None,
        )
        if selected_leader:
            st.session_state["leaders"] = selected_leader["leaders"]
            st.session_state["leader_phone"] = selected_leader["leader_phone"]
            st.session_state["leader_clipboard_message"] = ("success", "已填入选中的领队信息。")

    def save_current_leader() -> None:
        current_leaders = str(st.session_state.get("leaders") or "").strip()
        current_phone = str(st.session_state.get("leader_phone") or "").strip()
        if not current_leaders or not current_phone:
            st.session_state["leader_clipboard_message"] = ("warning", "领队和领队电话都要填写后才能保存。")
            return

        leaders = normalize_leader_clipboard(st.session_state.get("leader_clipboard", []))
        current_leader = {"leaders": current_leaders, "leader_phone": current_phone}
        current_label = leader_clipboard_label(current_leader)
        if any(leader_clipboard_label(leader) == current_label for leader in leaders):
            st.session_state["leader_clipboard_selected"] = current_label
            st.session_state["leader_clipboard_message"] = ("info", "这组领队信息已经在领队库里。")
            return

        leaders.append(current_leader)
        save_leader_clipboard(leaders)
        st.session_state["leader_clipboard"] = leaders
        st.session_state["leader_clipboard_selected"] = current_label
        st.session_state["leader_clipboard_message"] = ("success", "已保存当前领队信息。")

    def delete_selected_leader() -> None:
        selected_label = str(st.session_state.get("leader_clipboard_selected") or "").strip()
        if not selected_label:
            st.session_state["leader_clipboard_message"] = ("warning", "请先选择要删除的领队信息。")
            return

        leaders = [
            leader
            for leader in normalize_leader_clipboard(st.session_state.get("leader_clipboard", []))
            if leader_clipboard_label(leader) != selected_label
        ]
        save_leader_clipboard(leaders)
        st.session_state["leader_clipboard"] = leaders
        if leaders:
            st.session_state["leader_clipboard_selected"] = leader_clipboard_label(leaders[0])
        else:
            st.session_state.pop("leader_clipboard_selected", None)
        st.session_state["leader_clipboard_message"] = ("success", "已删除选中的领队信息。")

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
        key="route",
        placeholder="例如：西宁-祁连-七彩丹霞-敦煌-青海湖-西宁",
    )
    routes = normalize_route_clipboard(st.session_state.get("route_clipboard", []))
    st.session_state["route_clipboard"] = routes
    if routes and st.session_state.get("route_clipboard_selected") not in routes:
        st.session_state["route_clipboard_selected"] = routes[0]
    if routes:
        st.selectbox("路线库", routes, key="route_clipboard_selected")
        route_action_col1, route_action_col2, route_action_col3 = st.columns(3)
        with route_action_col1:
            st.button("填入路线", on_click=use_selected_route, use_container_width=True)
        with route_action_col2:
            st.button("保存当前路线", on_click=save_current_route, use_container_width=True)
        with route_action_col3:
            st.button("删除选中路线", on_click=delete_selected_route, use_container_width=True)
    else:
        st.markdown('<div class="field-note">路线库为空，保存当前路线后下次可以直接选择填入。</div>', unsafe_allow_html=True)
        st.button("保存当前路线", on_click=save_current_route, use_container_width=True)

    if "route_clipboard_message" in st.session_state:
        message_type, message_text = st.session_state.pop("route_clipboard_message")
        if message_type == "success":
            st.success(message_text)
        elif message_type == "warning":
            st.warning(message_text)
        else:
            st.info(message_text)

    col1, col2 = st.columns(2)
    with col1:
        leaders = st.text_input(
            "领队",
            key="leaders",
            placeholder="多个领队用顿号分隔",
        )
    with col2:
        leader_phone = st.text_input(
            "领队电话",
            key="leader_phone",
            placeholder="填写手机号",
        )
    leader_entries = normalize_leader_clipboard(st.session_state.get("leader_clipboard", []))
    st.session_state["leader_clipboard"] = leader_entries
    leader_labels = [leader_clipboard_label(leader) for leader in leader_entries]
    if leader_labels and st.session_state.get("leader_clipboard_selected") not in leader_labels:
        st.session_state["leader_clipboard_selected"] = leader_labels[0]
    if leader_labels:
        st.selectbox("领队库", leader_labels, key="leader_clipboard_selected")
        leader_action_col1, leader_action_col2, leader_action_col3 = st.columns(3)
        with leader_action_col1:
            st.button("填入领队", on_click=use_selected_leader, use_container_width=True)
        with leader_action_col2:
            st.button("保存当前领队", on_click=save_current_leader, use_container_width=True)
        with leader_action_col3:
            st.button("删除选中领队", on_click=delete_selected_leader, use_container_width=True)
    else:
        st.markdown('<div class="field-note">领队库为空，保存当前领队后下次可以直接选择填入。</div>', unsafe_allow_html=True)
        st.button("保存当前领队", on_click=save_current_leader, use_container_width=True)

    if "leader_clipboard_message" in st.session_state:
        message_type, message_text = st.session_state.pop("leader_clipboard_message")
        if message_type == "success":
            st.success(message_text)
        elif message_type == "warning":
            st.warning(message_text)
        else:
            st.info(message_text)

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
