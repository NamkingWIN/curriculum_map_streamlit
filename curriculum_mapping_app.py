"""
Ứng dụng Curriculum Mapping dùng Streamlit.

Mục đích:
- Kết hợp 3 mã nguồn hiện có:
  1) draw_flow.py: vẽ sơ đồ tiến trình/quan hệ học trước - tiên quyết
  2) draw_flow_plo.py: vẽ sơ đồ tiến trình theo PLO
  3) progress_check.py: kiểm tra logic dữ liệu quan hệ học phần
- Hiển thị dữ liệu động từ file Excel database chuẩn hóa.

Cách chạy:
    pip install streamlit pandas openpyxl
    streamlit run curriculum_mapping_app.py

Quy ước:
- Đặt file này cùng thư mục với draw_flow.py, draw_flow_plo.py, progress_check.py
  và progress_database_normalized.xlsx.
- Có thể dùng file mặc định hoặc upload file .xlsx/.xlsm trực tiếp trên giao diện.
"""

from __future__ import annotations

import io
import tempfile
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Import 3 file code hiện có. Các file này phải nằm cùng thư mục với app.
import draw_flow
import draw_flow_plo
import progress_check


APP_TITLE = "Curriculum Mapping"
DEFAULT_INPUT = "progress_database_normalized.xlsx"


# =========================
# TIỆN ÍCH CHUNG
# =========================

@st.cache_data(show_spinner=False)
def read_workbook_bytes(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    """Đọc toàn bộ workbook để preview dữ liệu trong app."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    return {
        sheet: pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, dtype=object)
        for sheet in xls.sheet_names
    }


def get_input_file() -> tuple[Path, bytes, str]:
    """Lấy file đầu vào từ uploader hoặc file mặc định."""
    uploaded = st.sidebar.file_uploader(
        "Chọn file database chuẩn hóa (.xlsx/.xlsm)",
        type=["xlsx", "xlsm"],
    )

    if uploaded is not None:
        suffix = Path(uploaded.name).suffix or ".xlsx"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(uploaded.getvalue())
        tmp.close()
        return Path(tmp.name), uploaded.getvalue(), uploaded.name

    default_path = Path(DEFAULT_INPUT)
    if not default_path.exists():
        st.error(
            f"Không tìm thấy file mặc định '{DEFAULT_INPUT}'. "
            "Hãy upload file database chuẩn hóa ở thanh bên trái."
        )
        st.stop()

    data = default_path.read_bytes()
    return default_path, data, default_path.name


def dataframe_download_button(df: pd.DataFrame, file_name: str, label: str):
    """Tạo nút tải Excel từ một DataFrame."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    st.download_button(
        label=label,
        data=buffer.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def workbook_download_button(sheets: dict[str, pd.DataFrame], file_name: str, label: str):
    """Tạo nút tải Excel gồm nhiều sheet."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
    st.download_button(
        label=label,
        data=buffer.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def show_html_dynamic(html_content: str, height: int = 760):
    """Hiển thị HTML tương tác trong Streamlit."""
    components.html(html_content, height=height, scrolling=True)


# =========================
# TAB 1: TỔNG QUAN DATABASE
# =========================

def page_database_preview(file_bytes: bytes):
    st.subheader("Dữ liệu database")
    workbook = read_workbook_bytes(file_bytes)

    sheet_names = list(workbook.keys())
    selected_sheet = st.selectbox("Chọn sheet để xem", sheet_names)
    df = workbook[selected_sheet]

    c1, c2, c3 = st.columns(3)
    c1.metric("Số sheet", len(sheet_names))
    c2.metric("Số dòng sheet đang xem", len(df))
    c3.metric("Số cột sheet đang xem", len(df.columns))

    st.dataframe(df, use_container_width=True, height=420)


# =========================
# TAB 2: KIỂM TRA LOGIC
# =========================

def run_progress_check(input_file: Path) -> dict[str, Any]:
    course_df, relation_type_df, course_relation_df = progress_check.read_database(input_file)
    courses = progress_check.build_course_dictionary(course_df)
    relation_types = progress_check.build_relation_type_lookup(relation_type_df)
    results, all_relations, missing_codes, redundant_conditions = progress_check.check_relations(
        course_relation_df=course_relation_df,
        courses=courses,
        relation_types=relation_types,
    )

    result_df = pd.DataFrame(results)
    all_relations_df = pd.DataFrame(all_relations)
    missing_df = pd.DataFrame(missing_codes)
    redundant_df = pd.DataFrame(redundant_conditions)
    course_out_df = pd.DataFrame(courses.values())

    def count_level(level: str) -> int:
        if result_df.empty or "Mức độ" not in result_df.columns:
            return 0
        return int((result_df["Mức độ"] == level).sum())

    summary_df = pd.DataFrame([
        ["Tổng số học phần", len(courses)],
        ["Tổng số quan hệ điều kiện", len(all_relations)],
        ["Tổng số vấn đề cần xem lại", len(results)],
        ["Mâu thuẫn nặng", count_level("Mâu thuẫn nặng")],
        ["Cùng học kỳ - cần kiểm tra", count_level("Cùng học kỳ - cần kiểm tra")],
        ["Mã không tìm thấy", count_level("Mã không tìm thấy")],
        ["Không đủ dữ liệu học kỳ", count_level("Không đủ dữ liệu học kỳ")],
        ["Điều kiện có thể bị dư", count_level("Điều kiện có thể bị dư")],
    ], columns=["Nội dung", "Giá trị"])

    return {
        "summary": summary_df,
        "results": result_df,
        "all_relations": all_relations_df,
        "missing": missing_df,
        "redundant": redundant_df,
        "courses": course_out_df,
    }


def page_progress_check(input_file: Path):
    st.subheader("Kiểm tra logic quan hệ học phần")

    try:
        data = run_progress_check(input_file)
    except Exception as exc:
        st.error(f"Không kiểm tra được dữ liệu: {exc}")
        return

    summary = data["summary"]
    values = dict(zip(summary["Nội dung"], summary["Giá trị"]))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Học phần", values.get("Tổng số học phần", 0))
    c2.metric("Quan hệ", values.get("Tổng số quan hệ điều kiện", 0))
    c3.metric("Vấn đề", values.get("Tổng số vấn đề cần xem lại", 0))
    c4.metric("Điều kiện có thể bị dư", values.get("Điều kiện có thể bị dư", 0))

    st.markdown("#### Tổng hợp")
    st.dataframe(summary, use_container_width=True, hide_index=True)

    st.markdown("#### Kết quả kiểm tra")
    if data["results"].empty:
        st.success("Không phát hiện vấn đề cần xem lại.")
    else:
        levels = sorted(data["results"]["Mức độ"].dropna().unique()) if "Mức độ" in data["results"].columns else []
        selected_levels = st.multiselect(
            "Lọc theo mức độ",
            levels,
            default=levels,
        )
        view_df = data["results"]
        if selected_levels and "Mức độ" in view_df.columns:
            view_df = view_df[view_df["Mức độ"].isin(selected_levels)]
        st.dataframe(view_df, use_container_width=True, height=420)

    with st.expander("Điều kiện có thể bị dư"):
        st.dataframe(data["redundant"], use_container_width=True, height=300)

    with st.expander("Đối chiếu tất cả quan hệ"):
        st.dataframe(data["all_relations"], use_container_width=True, height=300)

    workbook_download_button(
        {
            "Tong hop": data["summary"],
            "Ket qua kiem tra": data["results"],
            "Dieu kien co the bi du": data["redundant"],
            "Ma khong tim thay": data["missing"],
            "Doi chieu tat ca": data["all_relations"],
            "Danh muc hoc phan": data["courses"],
        },
        "KET_QUA_KIEM_TRA_APP.xlsx",
        "Tải kết quả kiểm tra Excel",
    )


# =========================
# TAB 3: SƠ ĐỒ TIẾN TRÌNH
# =========================

def build_flow_html(input_file: Path) -> tuple[str, int, int, int]:
    course_df = draw_flow.read_required_sheet(input_file, draw_flow.SHEET_COURSE)
    course_relation_df = draw_flow.read_required_sheet(input_file, draw_flow.SHEET_COURSE_RELATION)
    relation_type_df = draw_flow.read_required_sheet(input_file, draw_flow.SHEET_RELATION_TYPE)
    course_plo_df = draw_flow.read_optional_sheet(input_file, draw_flow.SHEET_COURSE_PLO)

    courses = draw_flow.build_course_data(course_df)
    relation_type_map = draw_flow.build_relation_type_map(relation_type_df)
    relations = draw_flow.build_relations(course_relation_df, relation_type_map, courses)
    draw_flow.add_plo_data(course_plo_df, courses)
    draw_flow.add_border_colors(courses)

    grouped = draw_flow.group_courses_by_semester(courses)
    positions, svg_width, svg_height = draw_flow.make_positions(grouped)
    svg_content = draw_flow.make_svg(courses, relations, grouped, positions, svg_width, svg_height)
    html_content = draw_flow.make_html(svg_content)
    return html_content, len(courses), len(relations), svg_height


def page_flow(input_file: Path):
    st.subheader("Sơ đồ tiến trình đào tạo")

    try:
        html_content, course_count, relation_count, svg_height = build_flow_html(input_file)
    except Exception as exc:
        st.error(f"Không tạo được sơ đồ tiến trình: {exc}")
        return

    c1, c2 = st.columns(2)
    c1.metric("Số học phần", course_count)
    c2.metric("Số quan hệ", relation_count)

    st.info("Click vào một môn để làm nổi các quan hệ liên quan. Rê chuột vào môn để xem tooltip.")
    show_html_dynamic(html_content, height=min(max(svg_height + 120, 680), 1100))

    st.download_button(
        "Tải HTML sơ đồ tiến trình",
        data=html_content.encode("utf-8"),
        file_name="So_do_tien_trinh_dao_tao_app.html",
        mime="text/html",
    )


# =========================
# TAB 4: SƠ ĐỒ PLO
# =========================

def build_plo_html(input_file: Path, selected_plos: list[str] | None) -> tuple[str, int, int, list[str], int]:
    config = draw_flow_plo.read_config(input_file)
    courses, final_plos, _sheet, warnings = draw_flow_plo.read_courses(input_file, config, selected_plos)
    grouped = draw_flow_plo.group_courses_by_semester(courses)
    positions, width, height = draw_flow_plo.make_positions(grouped)
    svg = draw_flow_plo.make_svg(courses, grouped, positions, width, height, final_plos)
    html_content = draw_flow_plo.make_html(svg, final_plos)
    mapped_courses = sum(1 for c in courses.values() if c.get("plo_map"))
    return html_content, len(courses), mapped_courses, warnings, height


def get_all_plos(input_file: Path) -> list[str]:
    try:
        plo_df = pd.read_excel(input_file, sheet_name=draw_flow_plo.SHEET_PLO, dtype=str, keep_default_na=False)
        if "is_active" in plo_df.columns:
            plo_df = plo_df[plo_df["is_active"].map(lambda x: draw_flow_plo.yes_no(x) or draw_flow_plo.safe_str(x) == "1")]
        if "plo_no" in plo_df.columns:
            plo_df["_sort"] = plo_df["plo_no"].map(lambda x: draw_flow_plo.safe_int(x) if draw_flow_plo.safe_int(x) is not None else 9999)
            plo_df = plo_df.sort_values(["_sort", "plo_id"])
        return [draw_flow_plo.safe_str(x).upper() for x in plo_df["plo_id"].tolist() if draw_flow_plo.safe_str(x)]
    except Exception:
        return []


def page_plo(input_file: Path):
    st.subheader("Sơ đồ tiến trình đào tạo theo PLO")

    all_plos = get_all_plos(input_file)
    selected = st.multiselect(
        "Chọn PLO muốn hiển thị",
        all_plos,
        default=all_plos,
    )
    selected_plos = selected if selected else None

    try:
        html_content, course_count, mapped_courses, warnings, svg_height = build_plo_html(input_file, selected_plos)
    except Exception as exc:
        st.error(f"Không tạo được sơ đồ PLO: {exc}")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Số học phần", course_count)
    c2.metric("Học phần có mapping PLO", mapped_courses)
    c3.metric("Số PLO đang chọn", len(selected_plos or all_plos))

    if warnings:
        with st.expander("Cảnh báo khi đọc dữ liệu"):
            for warning in warnings:
                st.warning(warning)

    st.info("Bấm nút PLO trong sơ đồ để lọc môn theo PLO. Sau đó click môn để hiện mũi tên từ môn tiên quyết/học trước sang môn được chọn.")
    show_html_dynamic(html_content, height=min(max(svg_height + 140, 700), 1150))

    st.download_button(
        "Tải HTML sơ đồ PLO",
        data=html_content.encode("utf-8"),
        file_name="So_do_tien_trinh_dao_tao_PLO_app.html",
        mime="text/html",
    )


# =========================
# MAIN APP
# =========================

def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="📚", layout="wide")

    st.title("📚 Curriculum Mapping")
    st.caption("Ứng dụng tổng hợp: xem database, kiểm tra logic, vẽ sơ đồ tiến trình và sơ đồ PLO từ file Excel chuẩn hóa.")

    input_file, file_bytes, file_name = get_input_file()
    st.sidebar.success(f"Đang dùng file: {file_name}")

    tab_preview, tab_check, tab_flow, tab_plo = st.tabs([
        "Database",
        "Kiểm tra logic",
        "Sơ đồ tiến trình",
        "Sơ đồ PLO",
    ])

    with tab_preview:
        page_database_preview(file_bytes)

    with tab_check:
        page_progress_check(input_file)

    with tab_flow:
        page_flow(input_file)

    with tab_plo:
        page_plo(input_file)


if __name__ == "__main__":
    main()
