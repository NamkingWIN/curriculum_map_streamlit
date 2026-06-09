import sys
import html
from pathlib import Path

import pandas as pd


INPUT_DEFAULT = "progress_database_normalized.xlsx"

SHEET_COURSE = "course"
SHEET_COURSE_RELATION = "course_relation"
SHEET_RELATION_TYPE = "relation"
SHEET_COURSE_PLO = "course_plo"

OUTPUT_HTML = "So_do_tien_trinh_dao_tao_database_20260609.html"
OUTPUT_SVG = "So_do_tien_trinh_dao_tao_database_20260609.svg"


# =========================
# MÀU SẮC
# =========================

COLOR_PHAI_DAT_TRUOC = "#f28c28"   # Cam: tiên quyết / phải đạt trước
COLOR_HOC_TRUOC = "#1f77b4"        # Xanh: học trước
COLOR_NORMAL_BORDER = "#b7c7d9"

BG_PHAI_DAT_TRUOC = "#fff4e6"
BG_HOC_TRUOC = "#eaf4ff"


# =========================
# HÀM TIỆN ÍCH
# =========================

def safe_str(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_course_code(value):
    """Chuẩn hóa mã học phần khi đọc từ database Excel."""
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if not text:
        return ""

    try:
        number = float(text)
        if number.is_integer():
            text = str(int(number))
    except Exception:
        pass

    return text.upper().replace(" ", "")


def is_truthy(value):
    """Đọc các cột cờ 1/0, TRUE/FALSE trong database."""
    text = safe_str(value).lower()
    return text in ["1", "1.0", "true", "yes", "y", "có", "co"]


def safe_int(value):
    if pd.isna(value):
        return None
    try:
        return int(float(value))
    except Exception:
        return None


def esc(value):
    return html.escape(str(value), quote=True)



def normalize_relation_type(value):
    """Chuẩn hóa tên loại quan hệ lấy từ sheet database relation."""
    text = safe_str(value).lower()

    if "học trước" in text or "hoc truoc" in text:
        return "Học trước"

    if "tiên quyết" in text or "tien quyet" in text or "phải đạt" in text or "phai dat" in text:
        return "Phải đạt trước"

    return safe_str(value)


def relation_color(level, relation_type):
    """
    Màu chỉ thể hiện loại quan hệ:
    - Phải đạt trước / tiên quyết: cam
    - Học trước: xanh
    """
    relation_type = safe_str(relation_type)

    if relation_type == "Phải đạt trước":
        return COLOR_PHAI_DAT_TRUOC

    if relation_type == "Học trước":
        return COLOR_HOC_TRUOC

    return "#999999"


def relation_dash(level):
    """
    Kiểu nét thể hiện tình trạng sắp xếp:
    - Mâu thuẫn nặng: chấm gạch / đứt khúc rõ
    - Cùng học kỳ: nét đứt nhẹ
    - Bình thường: nét liền
    """
    level = safe_str(level)

    if level == "Mâu thuẫn nặng":
        return "10 5 2 5"

    if level == "Cùng học kỳ - cần kiểm tra":
        return "6 4"

    return ""


def relation_status_text(level):
    """
    Nội dung tình trạng để hiển thị trong tooltip.
    """
    level = safe_str(level)

    if level == "Mâu thuẫn nặng":
        return "Mâu thuẫn: môn điều kiện bị xếp sau môn đang xét"

    if level == "Cùng học kỳ - cần kiểm tra":
        return "Cùng học kỳ: cần kiểm tra lại quy định học trước/tiên quyết"

    if level == "Không đủ dữ liệu học kỳ":
        return "Không đủ dữ liệu học kỳ để so sánh"

    return "Hợp lệ"


# =========================
# ĐỌC DỮ LIỆU
# =========================

def build_course_data(course_df):
    """
    Đọc danh mục học phần từ sheet database "course".
    Mỗi dòng trong sheet course là một học phần duy nhất.
    """
    required_columns = ["course_id", "course_name_vi", "semester"]
    missing = [col for col in required_columns if col not in course_df.columns]
    if missing:
        raise ValueError("Sheet course thiếu cột bắt buộc: " + ", ".join(missing))

    courses = {}

    for _, row in course_df.iterrows():
        code = normalize_course_code(row.get("course_id"))
        if not code:
            continue

        # Nếu có cột is_active thì chỉ vẽ học phần đang hoạt động.
        if "is_active" in course_df.columns and not is_truthy(row.get("is_active")):
            continue

        name = safe_str(row.get("course_name_vi"))
        semester = safe_int(row.get("semester"))
        display_order = safe_int(row.get("display_order"))
        credits = safe_int(row.get("credits"))
        course_type = safe_str(row.get("course_type"))
        outside_curriculum = 1 if is_truthy(row.get("outside_curriculum")) else 0

        courses[code] = {
            "code": code,
            "name": name,
            "semester": semester,
            "display_order": display_order,
            "credits": credits,
            "course_type": course_type,
            "outside_curriculum": outside_curriculum,
            "plo_levels": [],
            "relations_in": [],
            "relations_out": [],
            "border_colors": [],
        }

    return courses


def build_relation_type_map(relation_type_df):
    """
    Đọc sheet database "relation" để ánh xạ RT01/RT02 sang tên quan hệ.
    """
    if relation_type_df is None or relation_type_df.empty:
        return {}

    required_columns = ["relation_type_id", "relation_type_name"]
    missing = [col for col in required_columns if col not in relation_type_df.columns]
    if missing:
        raise ValueError("Sheet relation thiếu cột bắt buộc: " + ", ".join(missing))

    relation_type_map = {}
    for _, row in relation_type_df.iterrows():
        relation_type_id = safe_str(row.get("relation_type_id"))
        relation_type_name = normalize_relation_type(row.get("relation_type_name"))
        if relation_type_id:
            relation_type_map[relation_type_id] = relation_type_name

    return relation_type_map


def infer_relation_level_from_courses(source, target, courses):
    """
    Suy luận tình trạng sắp xếp học kỳ dựa trên học kỳ của môn điều kiện và môn đang xét.
    """
    source_semester = courses[source].get("semester")
    target_semester = courses[target].get("semester")

    if source_semester is None or target_semester is None:
        return "Không đủ dữ liệu học kỳ"
    if source_semester > target_semester:
        return "Mâu thuẫn nặng"
    if source_semester == target_semester:
        return "Cùng học kỳ - cần kiểm tra"
    return "Không mâu thuẫn"


def build_relations(course_relation_df, relation_type_map, courses):
    """
    Đọc quan hệ học phần từ sheet database "course_relation".

    Quy ước database:
    - source_course_id: môn điều kiện
    - target_course_id: môn đang xét
    Nghĩa là source_course_id → target_course_id.
    """
    required_columns = ["target_course_id", "source_course_id", "relation_type_id"]
    missing = [col for col in required_columns if col not in course_relation_df.columns]
    if missing:
        raise ValueError("Sheet course_relation thiếu cột bắt buộc: " + ", ".join(missing))

    relations = []

    for idx, row in course_relation_df.iterrows():
        target = normalize_course_code(row.get("target_course_id"))
        source = normalize_course_code(row.get("source_course_id"))

        if not source or not target:
            continue

        # Nếu có cột is_active thì chỉ vẽ quan hệ đang hoạt động.
        if "is_active" in course_relation_df.columns and not is_truthy(row.get("is_active")):
            continue

        # Chỉ vẽ quan hệ khi cả 2 mã học phần đều có trong sheet course.
        if source not in courses or target not in courses:
            continue

        relation_type_id = safe_str(row.get("relation_type_id"))
        relation_type = relation_type_map.get(relation_type_id, relation_type_id)
        relation_type = normalize_relation_type(relation_type)

        level = infer_relation_level_from_courses(source, target, courses)
        color = relation_color(level, relation_type)
        dash = relation_dash(level)
        relation_id = safe_str(row.get("relation_id")) or f"r{idx}"

        relation = {
            "id": f"r{idx}",
            "relation_id": relation_id,
            "source": source,
            "target": target,
            "type": relation_type,
            "level": level,
            "color": color,
            "dash": dash,
        }

        relations.append(relation)
        courses[source]["relations_out"].append(relation)
        courses[target]["relations_in"].append(relation)

    return relations


def add_plo_data(course_plo_df, courses):
    """
    Bổ sung thông tin PLO từ sheet database "course_plo" để hiển thị trong tooltip.
    """
    if course_plo_df is None or course_plo_df.empty:
        return

    required_columns = ["course_id", "plo_id"]
    missing = [col for col in required_columns if col not in course_plo_df.columns]
    if missing:
        raise ValueError("Sheet course_plo thiếu cột bắt buộc: " + ", ".join(missing))

    for _, row in course_plo_df.iterrows():
        code = normalize_course_code(row.get("course_id"))
        if code not in courses:
            continue

        if "is_mapped" in course_plo_df.columns and not is_truthy(row.get("is_mapped")):
            continue

        plo_id = safe_str(row.get("plo_id"))
        level_code = safe_str(row.get("level_code"))
        assessment = is_truthy(row.get("assessment"))

        if not plo_id:
            continue

        item = plo_id
        if level_code:
            item += f"-{level_code}"
        if assessment:
            item += "*"

        courses[code]["plo_levels"].append(item)


def add_border_colors(courses):
    """
    Viền ô học phần chỉ thể hiện vai trò "là điều kiện cho môn khác".

    Chỉ xét quan hệ đi ra của môn học, tức là course["relations_out"].

    - Viền cam: môn này là môn tiên quyết / phải đạt trước của môn khác.
    - Viền xanh: môn này là môn học trước của môn khác.
    - Không có quan hệ đi ra: viền xám bình thường, dù môn này có thể cần học
      trước / tiên quyết từ môn khác.
    """
    for code, course in courses.items():
        colors = []

        outgoing_relations = course["relations_out"]

        is_phai_dat_cho_mon_khac = any(
            r["type"] == "Phải đạt trước" for r in outgoing_relations
        )

        is_hoc_truoc_cho_mon_khac = any(
            r["type"] == "Học trước" for r in outgoing_relations
        )

        if is_phai_dat_cho_mon_khac:
            colors.append(COLOR_PHAI_DAT_TRUOC)

        if is_hoc_truoc_cho_mon_khac:
            colors.append(COLOR_HOC_TRUOC)

        course["border_colors"] = colors

def group_courses_by_semester(courses):
    """
    Nhóm học phần theo học kỳ.
    """
    grouped = {}

    for code, course in courses.items():
        semester = course["semester"]

        if semester is None:
            semester = 0

        grouped.setdefault(semester, []).append(course)

    for semester in grouped:
        grouped[semester].sort(key=lambda x: (x.get("display_order") is None, x.get("display_order") or 999999, x["name"], x["code"]))

    return dict(sorted(grouped.items(), key=lambda x: x[0]))


# =========================
# BỐ TRÍ SƠ ĐỒ
# =========================

def make_positions(grouped):
    """
    Tính vị trí từng ô học phần trong sơ đồ.
    """
    positions = {}

    margin_x = 80
    margin_y = 100

    col_width = 290
    card_width = 230
    card_height = 76
    row_gap = 34

    semesters = list(grouped.keys())

    for col_index, semester in enumerate(semesters):
        x = margin_x + col_index * col_width

        for row_index, course in enumerate(grouped[semester]):
            y = margin_y + row_index * (card_height + row_gap)

            positions[course["code"]] = {
                "x": x,
                "y": y,
                "w": card_width,
                "h": card_height,
                "semester": semester,
            }

    max_rows = max(len(items) for items in grouped.values()) if grouped else 1

    svg_width = margin_x * 2 + len(semesters) * col_width
    svg_height = margin_y + max_rows * (card_height + row_gap) + 180

    return positions, svg_width, svg_height


def edge_path(source_pos, target_pos):
    """
    Tạo đường cong từ môn điều kiện sang môn được học.
    """
    sx = source_pos["x"] + source_pos["w"]
    sy = source_pos["y"] + source_pos["h"] / 2

    tx = target_pos["x"]
    ty = target_pos["y"] + target_pos["h"] / 2

    # Nếu quan hệ ngược chiều hoặc cùng cột, cho đường đi vòng để dễ nhìn.
    if tx <= sx:
        offset = 55
        p1x = sx + offset
        p2x = tx - offset

        if ty >= sy:
            mid_y = max(sy, ty) + 44
        else:
            mid_y = min(sy, ty) - 44

        return (
            f"M {sx} {sy} "
            f"C {p1x} {sy}, {p1x} {mid_y}, {(sx + tx) / 2} {mid_y} "
            f"S {p2x} {ty}, {tx} {ty}"
        )

    dx = max(80, (tx - sx) / 2)

    return f"M {sx} {sy} C {sx + dx} {sy}, {tx - dx} {ty}, {tx} {ty}"


# =========================
# HIỂN THỊ Ô HỌC PHẦN
# =========================

def border_rects(course, pos):
    """
    Vẽ nhiều lớp viền màu cho ô học phần.
    """
    colors = course.get("border_colors", [])

    if not colors:
        colors = [COLOR_NORMAL_BORDER]

    rects = []

    for i, color in enumerate(colors):
        inset = i * 4
        x = pos["x"] + inset
        y = pos["y"] + inset
        w = pos["w"] - inset * 2
        h = pos["h"] - inset * 2

        rects.append(
            f'''
            <rect
                x="{x}"
                y="{y}"
                width="{w}"
                height="{h}"
                rx="10"
                ry="10"
                fill="none"
                stroke="{color}"
                stroke-width="2.5"
            />
            '''
        )

    return "\n".join(rects)


def wrap_text(text, max_chars=30):
    """
    Tách tên học phần thành nhiều dòng ngắn để hiển thị trong ô.
    """
    words = text.split()
    lines = []
    current = ""

    for word in words:
        if len(current + " " + word) <= max_chars:
            current = (current + " " + word).strip()
        else:
            if current:
                lines.append(current)
            current = word

    if current:
        lines.append(current)

    return lines[:3]


def course_bg(course):
    """
    Nền chỉ tô màu khi môn này là điều kiện cho môn khác.
    Chỉ xét quan hệ đi ra.
    """
    outgoing_relations = course["relations_out"]

    is_phai_dat_cho_mon_khac = any(
        r["type"] == "Phải đạt trước" for r in outgoing_relations
    )

    is_hoc_truoc_cho_mon_khac = any(
        r["type"] == "Học trước" for r in outgoing_relations
    )

    if is_phai_dat_cho_mon_khac and is_hoc_truoc_cho_mon_khac:
        return "#f8fbff"

    if is_phai_dat_cho_mon_khac:
        return BG_PHAI_DAT_TRUOC

    if is_hoc_truoc_cho_mon_khac:
        return BG_HOC_TRUOC

    return "#ffffff"


# =========================
# TOOLTIP CHO Ô HỌC PHẦN
# =========================

def course_relation_label(course, courses):
    """
    Tạo tooltip khi rê chuột vào ô học phần.

    Hiển thị rõ cả 2 chiều:
    1. Môn này là điều kiện cho môn khác.
    2. Môn này cần môn khác làm điều kiện.
    """
    lines = []

    code = course["code"]
    name = course["name"]

    lines.append(f"{code} - {name}")
    lines.append("")

    outgoing = course["relations_out"]
    incoming = course["relations_in"]

    if outgoing:
        lines.append("Môn này là học trước / tiên quyết của:")

        for r in outgoing:
            source = r["source"]
            target = r["target"]

            source_name = courses[source]["name"] if source in courses else ""
            target_name = courses[target]["name"] if target in courses else ""

            relation_type = safe_str(r["type"])
            status = relation_status_text(r["level"])

            lines.append(
                f"- {source} - {source_name} → {target} - {target_name}"
            )
            lines.append(
                f"  Loại: {relation_type}; Tình trạng: {status}"
            )

        lines.append("")

    if incoming:
        lines.append("Môn này cần học trước / cần đạt trước:")

        for r in incoming:
            source = r["source"]
            target = r["target"]

            source_name = courses[source]["name"] if source in courses else ""
            target_name = courses[target]["name"] if target in courses else ""

            relation_type = safe_str(r["type"])
            status = relation_status_text(r["level"])

            lines.append(
                f"- {source} - {source_name} → {target} - {target_name}"
            )
            lines.append(
                f"  Loại: {relation_type}; Tình trạng: {status}"
            )

        lines.append("")

    if not outgoing and not incoming:
        lines.append("Môn học này chưa có quan hệ học trước / tiên quyết trong dữ liệu.")

    return "\n".join(lines)


def course_hover_label(course, courses):
    """
    Tooltip ngắn gọn khi rê chuột vào ô học phần.

    Nội dung chỉ tóm tắt các quan hệ theo chiều:
    - Có môn học tiên quyết/học trước: các môn là điều kiện của môn này.
    - Là môn học tiên quyết/học trước của: các môn mà môn này là điều kiện.
    """
    code = course["code"]
    name = course["name"]
    semester = course.get("semester")

    if semester is None or semester == 0:
        semester_text = "Chưa rõ học kỳ"
    else:
        semester_text = f"Học kỳ {semester}"

    def unique_codes(relations, relation_type, field):
        values = []
        seen = set()

        for r in relations:
            if r["type"] != relation_type:
                continue

            value = r[field]
            if value and value not in seen:
                values.append(value)
                seen.add(value)

        return values

    incoming = course["relations_in"]
    outgoing = course["relations_out"]

    co_tien_quyet = unique_codes(incoming, "Phải đạt trước", "source")
    la_tien_quyet_cua = unique_codes(outgoing, "Phải đạt trước", "target")

    co_hoc_truoc = unique_codes(incoming, "Học trước", "source")
    la_hoc_truoc_cua = unique_codes(outgoing, "Học trước", "target")

    lines = [f"{code} - {name}", semester_text]

    plo_levels = course.get("plo_levels", [])
    if plo_levels:
        lines.append("PLO: " + ", ".join(plo_levels))

    if co_tien_quyet:
        lines.append("Có môn học tiên quyết: " + ", ".join(co_tien_quyet))

    if la_tien_quyet_cua:
        lines.append("Là môn học tiên quyết của: " + ", ".join(la_tien_quyet_cua))

    if co_hoc_truoc:
        lines.append("Có môn học trước: " + ", ".join(co_hoc_truoc))

    if la_hoc_truoc_cua:
        lines.append("Là môn học trước của: " + ", ".join(la_hoc_truoc_cua))

    if len(lines) == 2:
        lines.append("Không có quan hệ tiên quyết/học trước.")

    lines.append("Click để làm nổi các quan hệ liên quan.")

    return "\n".join(lines)


# =========================
# TẠO SVG
# =========================

def make_svg(courses, relations, grouped, positions, svg_width, svg_height):
    """
    Tạo SVG sơ đồ.
    """
    semesters = list(grouped.keys())

    svg_parts = []

    svg_parts.append(f'''
<svg
    xmlns="http://www.w3.org/2000/svg"
    width="{svg_width}"
    height="{svg_height}"
    viewBox="0 0 {svg_width} {svg_height}"
>
<defs>
    <marker id="arrow-orange" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto">
        <path d="M0,0 L0,6 L9,3 z" fill="{COLOR_PHAI_DAT_TRUOC}" />
    </marker>

    <marker id="arrow-blue" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto">
        <path d="M0,0 L0,6 L9,3 z" fill="{COLOR_HOC_TRUOC}" />
    </marker>

    <marker id="arrow-gray" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto">
        <path d="M0,0 L0,6 L9,3 z" fill="#999999" />
    </marker>

    <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
        <feDropShadow dx="0" dy="2" stdDeviation="2" flood-color="#000000" flood-opacity="0.12"/>
    </filter>
</defs>

<style>
    .semester-title {{
        font-family: Arial, sans-serif;
        font-size: 18px;
        font-weight: bold;
        fill: #1f2937;
    }}

    .course-card {{
        filter: url(#shadow);
        cursor: pointer;
        transition: opacity 0.15s ease;
    }}

    .course-bg {{
        stroke: #e5e7eb;
        stroke-width: 1;
    }}

    .course-code {{
        font-family: Arial, sans-serif;
        font-size: 12px;
        font-weight: bold;
        fill: #374151;
    }}

    .course-name {{
        font-family: Arial, sans-serif;
        font-size: 11px;
        fill: #111827;
    }}

    .edge {{
        fill: none;
        stroke-width: 1.6;
        opacity: 0.38;
        transition: opacity 0.15s ease, stroke-width 0.15s ease;
        pointer-events: none;
    }}

    .edge.active {{
        stroke-width: 3.4;
        opacity: 1;
    }}

    .edge.dim {{
        opacity: 0.04;
    }}

    .course-card.dim {{
        opacity: 0.25;
    }}

    .legend-text {{
        font-family: Arial, sans-serif;
        font-size: 13px;
        fill: #374151;
    }}

    .legend-title {{
        font-family: Arial, sans-serif;
        font-size: 14px;
        font-weight: bold;
        fill: #111827;
    }}
</style>
''')

    # Nền
    svg_parts.append(
        f'<rect x="0" y="0" width="{svg_width}" height="{svg_height}" fill="#f8fafc" />'
    )

    # Tiêu đề học kỳ
    for col_index, semester in enumerate(semesters):
        x = 80 + col_index * 290
        title = "Chưa rõ HK" if semester == 0 else f"Học kỳ {semester}"

        svg_parts.append(
            f'''
            <text x="{x}" y="55" class="semester-title">{esc(title)}</text>
            <line x1="{x}" y1="70" x2="{x + 230}" y2="70" stroke="#cbd5e1" stroke-width="2"/>
            '''
        )

    # Mũi tên
    for r in relations:
        if r["source"] not in positions or r["target"] not in positions:
            continue

        source_pos = positions[r["source"]]
        target_pos = positions[r["target"]]

        path = edge_path(source_pos, target_pos)

        marker = "arrow-gray"

        if r["color"] == COLOR_PHAI_DAT_TRUOC:
            marker = "arrow-orange"
        elif r["color"] == COLOR_HOC_TRUOC:
            marker = "arrow-blue"

        dash_attr = f'stroke-dasharray="{r["dash"]}"' if r["dash"] else ""

        svg_parts.append(
            f'''
            <path
                id="{r["id"]}"
                class="edge"
                d="{path}"
                stroke="{r["color"]}"
                marker-end="url(#{marker})"
                {dash_attr}
            />
            '''
        )

    # Ô học phần
    for semester, items in grouped.items():
        for course in items:
            code = course["code"]
            pos = positions[code]
            bg = course_bg(course)

            name_lines = wrap_text(course["name"], max_chars=30)

            text_lines = []
            text_lines.append(
                f'<text x="{pos["x"] + 12}" y="{pos["y"] + 22}" class="course-code">{esc(code)}</text>'
            )

            start_y = pos["y"] + 42

            for i, line in enumerate(name_lines):
                text_lines.append(
                    f'<text x="{pos["x"] + 12}" y="{start_y + i * 14}" class="course-name">{esc(line)}</text>'
                )

            related_edges = [r["id"] for r in course["relations_in"] + course["relations_out"]]
            related_courses = set()

            for r in course["relations_in"] + course["relations_out"]:
                related_courses.add(r["source"])
                related_courses.add(r["target"])

            related_edges_attr = ",".join(related_edges)
            related_courses_attr = ",".join(related_courses)
            course_tooltip = course_hover_label(course, courses)

            svg_parts.append(
                f'''
                <g
                    id="course-{esc(code)}"
                    class="course-card"
                    data-code="{esc(code)}"
                    data-edges="{esc(related_edges_attr)}"
                    data-courses="{esc(related_courses_attr)}"
                    data-tooltip="{esc(course_tooltip)}"
                >
                    <rect
                        x="{pos["x"]}"
                        y="{pos["y"]}"
                        width="{pos["w"]}"
                        height="{pos["h"]}"
                        rx="10"
                        ry="10"
                        fill="{bg}"
                        class="course-bg"
                    />
                    {border_rects(course, pos)}
                    {"".join(text_lines)}
                </g>
                '''
            )

    # Chú giải
    legend_x = 80
    legend_y = svg_height - 120

    svg_parts.append(f'''
    <g id="legend">
        <text x="{legend_x}" y="{legend_y - 18}" class="legend-title">Chú giải</text>

        <line x1="{legend_x}" y1="{legend_y}" x2="{legend_x + 40}" y2="{legend_y}" stroke="{COLOR_PHAI_DAT_TRUOC}" stroke-width="3"/>
        <text x="{legend_x + 50}" y="{legend_y + 4}" class="legend-text">Cam: tiên quyết / phải đạt trước</text>

        <line x1="{legend_x + 300}" y1="{legend_y}" x2="{legend_x + 340}" y2="{legend_y}" stroke="{COLOR_HOC_TRUOC}" stroke-width="3"/>
        <text x="{legend_x + 350}" y="{legend_y + 4}" class="legend-text">Xanh: học trước</text>

        <line x1="{legend_x}" y1="{legend_y + 32}" x2="{legend_x + 40}" y2="{legend_y + 32}" stroke="#555555" stroke-width="3"/>
        <text x="{legend_x + 50}" y="{legend_y + 36}" class="legend-text">Nét liền: sắp xếp hợp lệ</text>

        <line x1="{legend_x + 300}" y1="{legend_y + 32}" x2="{legend_x + 340}" y2="{legend_y + 32}" stroke="#555555" stroke-width="3" stroke-dasharray="6 4"/>
        <text x="{legend_x + 350}" y="{legend_y + 36}" class="legend-text">Nét đứt: cùng học kỳ, cần kiểm tra</text>

        <line x1="{legend_x}" y1="{legend_y + 64}" x2="{legend_x + 40}" y2="{legend_y + 64}" stroke="#555555" stroke-width="3" stroke-dasharray="10 5 2 5"/>
        <text x="{legend_x + 50}" y="{legend_y + 68}" class="legend-text">Chấm gạch: môn điều kiện bị xếp sau môn đang xét</text>

        <rect x="{legend_x + 430}" y="{legend_y + 50}" width="44" height="28" rx="6" fill="white" stroke="{COLOR_PHAI_DAT_TRUOC}" stroke-width="2.5"/>
        <rect x="{legend_x + 434}" y="{legend_y + 54}" width="36" height="20" rx="5" fill="none" stroke="{COLOR_HOC_TRUOC}" stroke-width="2.5"/>
        <text x="{legend_x + 490}" y="{legend_y + 68}" class="legend-text">Ô có nhiều viền: môn có nhiều loại quan hệ</text>
    </g>
''')

    svg_parts.append("</svg>")

    return "\n".join(svg_parts)


# =========================
# TẠO HTML TƯƠNG TÁC
# =========================

def make_html(svg_content):
    """
    Bọc SVG vào HTML có tương tác click và tooltip khi rê chuột.
    - Click vào ô học phần: làm nổi các môn và mũi tên liên quan.
    - Rê chuột vào ô học phần: chỉ hiển thị thông tin ngắn gọn của học phần.
    """
    html_content = f"""
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<title>Sơ đồ tiến trình đào tạo</title>

<style>
    body {{
        margin: 0;
        font-family: Arial, sans-serif;
        background: #f8fafc;
        color: #111827;
    }}

    .topbar {{
        position: sticky;
        top: 0;
        z-index: 10;
        background: white;
        border-bottom: 1px solid #e5e7eb;
        padding: 12px 18px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }}

    .topbar h1 {{
        margin: 0 0 6px 0;
        font-size: 20px;
    }}

    .topbar p {{
        margin: 0;
        font-size: 13px;
        color: #4b5563;
    }}

    .canvas {{
        overflow: auto;
        height: calc(100vh - 82px);
        padding: 16px;
    }}

    svg {{
        background: #f8fafc;
    }}

    .cm-hover-tooltip {{
        position: fixed;
        display: none;
        max-width: 620px;
        max-height: 70vh;
        overflow: auto;
        background: #111827;
        color: white;
        padding: 10px 12px;
        border-radius: 8px;
        font-size: 13px;
        line-height: 1.45;
        z-index: 9999;
        pointer-events: none;
        white-space: pre-line;
        box-shadow: 0 8px 24px rgba(0,0,0,0.18);
    }}
</style>
</head>

<body>
<div class="topbar">
    <h1>Sơ đồ tiến trình đào tạo</h1>
    <p>
        Click vào khung môn học để làm nổi các môn và mũi tên liên quan.
        Rê chuột vào khung môn học để xem mã, tên, học kỳ và PLO.
    </p>
</div>

<div class="canvas">
{svg_content}
</div>

<div id="cm-hover-tooltip" class="cm-hover-tooltip"></div>

<script>
    const cards = document.querySelectorAll('.course-card');
    const edges = document.querySelectorAll('.edge');
    const tooltip = document.getElementById('cm-hover-tooltip');
    let selectedCard = null;

    function clearHighlight() {{
        selectedCard = null;

        cards.forEach(card => {{
            card.classList.remove('dim');
            card.classList.remove('selected');
        }});

        edges.forEach(edge => {{
            edge.classList.remove('active');
            edge.classList.remove('dim');
        }});

        hideTooltip();
    }}

    function highlightCard(card) {{
        const relatedEdges = card.dataset.edges
            ? card.dataset.edges.split(',').filter(x => x)
            : [];

        const relatedCourses = card.dataset.courses
            ? card.dataset.courses.split(',').filter(x => x)
            : [];

        cards.forEach(c => {{
            const code = c.dataset.code;

            if (!relatedCourses.includes(code)) {{
                c.classList.add('dim');
            }} else {{
                c.classList.remove('dim');
            }}
        }});

        edges.forEach(e => {{
            if (relatedEdges.includes(e.id)) {{
                e.classList.add('active');
                e.classList.remove('dim');
            }} else {{
                e.classList.remove('active');
                e.classList.add('dim');
            }}
        }});

        card.classList.remove('dim');
        card.classList.add('selected');
    }}

    function showTooltip(card, event) {{
        const text = card.dataset.tooltip || '';
        if (!text) return;

        tooltip.textContent = text;
        tooltip.style.display = 'block';
        moveTooltip(event);
    }}

    function hideTooltip() {{
        tooltip.style.display = 'none';
    }}

    function moveTooltip(event) {{
        const padding = 16;
        const offset = 14;

        tooltip.style.left = event.clientX + offset + 'px';
        tooltip.style.top = event.clientY + offset + 'px';

        const rect = tooltip.getBoundingClientRect();

        if (rect.right > window.innerWidth - padding) {{
            tooltip.style.left = event.clientX - rect.width - offset + 'px';
        }}

        if (rect.bottom > window.innerHeight - padding) {{
            tooltip.style.top = event.clientY - rect.height - offset + 'px';
        }}
    }}

    cards.forEach(card => {{
        card.addEventListener('mouseenter', (event) => {{
            showTooltip(card, event);
        }});

        card.addEventListener('mousemove', (event) => {{
            if (tooltip.style.display === 'block') {{
                moveTooltip(event);
            }}
        }});

        card.addEventListener('mouseleave', () => {{
            hideTooltip();
        }});

        card.addEventListener('click', (event) => {{
            event.stopPropagation();
            hideTooltip();

            if (selectedCard === card) {{
                clearHighlight();
                return;
            }}

            clearHighlight();
            selectedCard = card;
            highlightCard(card);
        }});
    }});

    document.addEventListener('click', () => {{
        clearHighlight();
    }});

    document.addEventListener('keydown', (event) => {{
        if (event.key === 'Escape') {{
            clearHighlight();
        }}
    }});
</script>

</body>
</html>
"""
    return html_content


def read_required_sheet(input_file, sheet_name):
    """Đọc một sheet bắt buộc trong file database."""
    xls = pd.ExcelFile(input_file)
    if sheet_name not in xls.sheet_names:
        raise ValueError(f"Không tìm thấy sheet bắt buộc: {sheet_name}")
    return pd.read_excel(input_file, sheet_name=sheet_name)


def read_optional_sheet(input_file, sheet_name):
    """Đọc một sheet không bắt buộc; nếu không có thì trả về None."""
    xls = pd.ExcelFile(input_file)
    if sheet_name not in xls.sheet_names:
        return None
    return pd.read_excel(input_file, sheet_name=sheet_name)


# =========================
# MAIN
# =========================

def main():
    if len(sys.argv) >= 2:
        input_file = Path(sys.argv[1])
    else:
        input_file = Path(INPUT_DEFAULT)

    if not input_file.exists():
        print(f"Không tìm thấy file: {input_file}")
        print("Cách dùng:")
        print('python draw_flow.py "progress_database_normalized.xlsx"')
        sys.exit(1)

    try:
        course_df = read_required_sheet(input_file, SHEET_COURSE)
        course_relation_df = read_required_sheet(input_file, SHEET_COURSE_RELATION)
        relation_type_df = read_required_sheet(input_file, SHEET_RELATION_TYPE)
        course_plo_df = read_optional_sheet(input_file, SHEET_COURSE_PLO)
    except Exception as e:
        print("Không đọc được file database hoặc thiếu sheet cần thiết.")
        print("File cần có tối thiểu các sheet: course, course_relation, relation.")
        print(str(e))
        sys.exit(1)

    courses = build_course_data(course_df)
    relation_type_map = build_relation_type_map(relation_type_df)
    relations = build_relations(course_relation_df, relation_type_map, courses)
    add_plo_data(course_plo_df, courses)

    add_border_colors(courses)

    grouped = group_courses_by_semester(courses)
    positions, svg_width, svg_height = make_positions(grouped)

    svg_content = make_svg(
        courses=courses,
        relations=relations,
        grouped=grouped,
        positions=positions,
        svg_width=svg_width,
        svg_height=svg_height,
    )

    html_content = make_html(svg_content)

    output_svg = input_file.with_name(OUTPUT_SVG)
    output_html = input_file.with_name(OUTPUT_HTML)

    output_svg.write_text(svg_content, encoding="utf-8")
    output_html.write_text(html_content, encoding="utf-8")

    print("Đã tạo sơ đồ từ database.")
    print(f"HTML: {output_html}")
    print(f"SVG : {output_svg}")
    print("")
    print(f"Số học phần: {len(courses)}")
    print(f"Số quan hệ: {len(relations)}")


if __name__ == "__main__":
    main()
