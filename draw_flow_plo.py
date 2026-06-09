import sys
import re
import html
from pathlib import Path
from collections import defaultdict

import pandas as pd

INPUT_DEFAULT = "progress_database_normalized.xlsx"
OUTPUT_HTML = "So_do_tien_trinh_dao_tao_PLO_database.html"
OUTPUT_SVG = "So_do_tien_trinh_dao_tao_PLO_database.svg"
CONFIG_SHEET = "CONFIG"

# Tên sheet trong file database đã chuẩn hóa
SHEET_COURSE = "course"
SHEET_PLO = "plo"
SHEET_COURSE_PLO = "course_plo"
SHEET_RELATION = "relation"
SHEET_COURSE_RELATION = "course_relation"

DEFAULT_CONFIG = {
    "DATA_SHEET": "",
    "HEADER_ROW": "1",
    "START_DATA_ROW": "2",
    "COL_MA_HP": "B",
    "COL_TEN_HP": "D",
    "COL_HOC_KY": "F",
    "COL_PHAI_DAT_TRUOC": "P",
    "COL_HOC_TRUOC": "Q",
    "SELECTED_PLOS": "",
    "ONLY_PLO_RELATED_COURSES": "NO",  # NO = vẫn thấy toàn bộ sơ đồ; bấm PLO sẽ lọc/làm nổi môn liên quan
}

PLO_COL_RE = re.compile(r"^PLO\d+$", re.IGNORECASE)
LEVEL_ORDER = {"I": 1, "R": 2, "M": 3}
LEVEL_LABEL = {"I": "Introduction", "R": "Reinforce", "M": "Mastery"}

# Màu viền theo yêu cầu
LEVEL_BORDER = {
    "I": "#22c55e",  # xanh lá
    "R": "#facc15",  # vàng
    "M": "#dc2626",  # đỏ
}

# Nền nhẹ theo mức để dễ nhìn khi chọn PLO
LEVEL_BG = {
    "I": "#dcfce7",  # xanh lá nền
    "R": "#fef3c7",  # vàng nền
    "M": "#fecaca",  # đỏ nền
}

NORMAL_BORDER = "#b7c7d9"
NORMAL_BG = "#ffffff"
ASSESS_COLOR = "#111827"
BADGE_TEXT_COLOR = "#78350f"  # nâu đậm cho dễ đọc trên các badge PLO
PREREQ_BORDER = "#f97316"
PREREQ_BG = "#fff7ed"
PREV_BORDER = "#2563eb"
PREV_BG = "#eff6ff"
BOTH_BORDER = "#7c3aed"
BOTH_BG = "#f5f3ff"


def safe_str(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def safe_int(value):
    text = safe_str(value)
    if not text:
        return None
    try:
        return int(float(text))
    except Exception:
        return None


def esc(value):
    return html.escape(str(value), quote=True)


def yes_no(value):
    return safe_str(value).upper() in {"YES", "Y", "TRUE", "1", "CÓ", "CO"}


def col_letter_to_index(letter: str) -> int:
    text = safe_str(letter).upper()
    if not re.fullmatch(r"[A-Z]+", text):
        raise ValueError(f"Mã cột không hợp lệ: {letter}")
    result = 0
    for char in text:
        result = result * 26 + ord(char) - ord("A") + 1
    return result - 1


def parse_semester(value):
    text = safe_str(value)
    if not text:
        return None
    m = re.search(r"\d+", text)
    return int(m.group()) if m else None


def normalize_code(value):
    text = safe_str(value).upper()
    text = re.sub(r"\s+", "", text)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def parse_course_codes(value) -> list[str]:
    """Tách danh sách mã học phần, hỗ trợ dấu phẩy, chấm phẩy và xuống dòng."""
    text = safe_str(value)
    if not text:
        return []
    codes = []
    seen = set()
    for part in re.split(r"[,;\n]+", text):
        code = normalize_code(part)
        if code and code not in seen:
            codes.append(code)
            seen.add(code)
    return codes


def read_config(file_path: Path) -> dict:
    config = DEFAULT_CONFIG.copy()
    try:
        cfg = pd.read_excel(file_path, sheet_name=CONFIG_SHEET, dtype=str, keep_default_na=False)
        for _, row in cfg.iterrows():
            key = safe_str(row.iloc[0])
            value = safe_str(row.iloc[1]) if len(row) > 1 else ""
            if key:
                config[key] = value
    except Exception:
        pass
    return config


def get_data_sheet_name(file_path: Path, config: dict) -> str:
    xls = pd.ExcelFile(file_path)
    wanted = safe_str(config.get("DATA_SHEET"))
    if wanted:
        if wanted not in xls.sheet_names:
            raise ValueError(f"Không tìm thấy sheet dữ liệu: {wanted}")
        return wanted
    for sheet in xls.sheet_names:
        if sheet != CONFIG_SHEET:
            return sheet
    raise ValueError("Không tìm thấy sheet dữ liệu.")


def column_name_from_config(df: pd.DataFrame, value: str) -> str:
    value = safe_str(value)
    if value in df.columns:
        return value
    if re.fullmatch(r"[A-Za-z]+", value):
        idx = col_letter_to_index(value)
        if idx < len(df.columns):
            return df.columns[idx]
    raise ValueError(f"Không xác định được cột: {value}")


def parse_plo_cell(value):
    """
    Trả về {'level': I/R/M, 'assess': bool, 'raw': text} hoặc None.
    Hỗ trợ I, R, M, I,A, R,A, M,A, M, A.
    Các giá trị không thuộc I/R/M/A, ví dụ 20, được bỏ qua.
    """
    raw = safe_str(value).upper()
    if not raw:
        return None
    tokens = [t.strip() for t in re.split(r"[,;\s]+", raw) if t.strip()]
    level = None
    assess = False
    for token in tokens:
        if token in LEVEL_ORDER and level is None:
            level = token
        elif token == "A":
            assess = True
    if level is None:
        return None
    return {"level": level, "assess": assess, "raw": raw}


def selected_plos_from_config(df: pd.DataFrame, config: dict, cli_plos: list[str] | None) -> list[str]:
    all_plos = [c for c in df.columns if PLO_COL_RE.match(str(c).strip())]
    if cli_plos:
        selected = cli_plos
    else:
        text = safe_str(config.get("SELECTED_PLOS"))
        selected = [x.strip() for x in re.split(r"[,;\n]+", text) if x.strip()] if text else all_plos
    missing = [p for p in selected if p not in df.columns]
    if missing:
        raise ValueError("Không tìm thấy các cột PLO: " + ", ".join(missing))
    return selected


def read_courses(file_path: Path, config: dict, cli_plos: list[str] | None):
    """
    Đọc dữ liệu từ file database đã chuẩn hóa.

    Cấu trúc kỳ vọng:
    - course: danh mục học phần
    - plo: danh mục PLO
    - course_plo: bảng mapping học phần - PLO
    - relation: danh mục loại quan hệ
    - course_relation: bảng quan hệ học phần
    """
    xls = pd.ExcelFile(file_path)
    required_sheets = [SHEET_COURSE, SHEET_PLO, SHEET_COURSE_PLO, SHEET_RELATION, SHEET_COURSE_RELATION]
    missing_sheets = [s for s in required_sheets if s not in xls.sheet_names]
    if missing_sheets:
        raise ValueError("File database thiếu sheet: " + ", ".join(missing_sheets))

    course_df = pd.read_excel(file_path, sheet_name=SHEET_COURSE, dtype=str, keep_default_na=False)
    plo_df = pd.read_excel(file_path, sheet_name=SHEET_PLO, dtype=str, keep_default_na=False)
    course_plo_df = pd.read_excel(file_path, sheet_name=SHEET_COURSE_PLO, dtype=str, keep_default_na=False)
    relation_df = pd.read_excel(file_path, sheet_name=SHEET_RELATION, dtype=str, keep_default_na=False)
    course_relation_df = pd.read_excel(file_path, sheet_name=SHEET_COURSE_RELATION, dtype=str, keep_default_na=False)

    for df in [course_df, plo_df, course_plo_df, relation_df, course_relation_df]:
        df.columns = [safe_str(c) for c in df.columns]

    warnings = []

    # Danh mục PLO đang hoạt động, sắp theo plo_no nếu có.
    if cli_plos:
        selected_plos = [safe_str(p).upper() for p in cli_plos if safe_str(p)]
    else:
        active_plo_df = plo_df.copy()
        if "is_active" in active_plo_df.columns:
            active_plo_df = active_plo_df[active_plo_df["is_active"].map(lambda x: yes_no(x) or safe_str(x) == "1")]
        active_plo_df["_sort"] = active_plo_df["plo_no"].map(lambda x: safe_int(x) if safe_int(x) is not None else 9999) if "plo_no" in active_plo_df.columns else 9999
        active_plo_df = active_plo_df.sort_values(["_sort", "plo_id"])
        selected_plos = [safe_str(x).upper() for x in active_plo_df["plo_id"].tolist() if safe_str(x)]

    if not selected_plos:
        selected_plos = sorted({safe_str(x).upper() for x in course_plo_df.get("plo_id", []) if safe_str(x)})

    selected_plo_set = set(selected_plos)

    # Đọc danh mục học phần.
    courses = {}
    for _, row in course_df.iterrows():
        if "is_active" in course_df.columns and not (yes_no(row.get("is_active")) or safe_str(row.get("is_active")) == "1"):
            continue

        code = normalize_code(row.get("course_id"))
        if not code:
            continue

        if code in courses:
            warnings.append(f"Trùng course_id {code}; giữ dòng đầu, bỏ qua dòng sau.")
            continue

        courses[code] = {
            "code": code,
            "name": safe_str(row.get("course_name_vi")),
            "semester": parse_semester(row.get("semester")),
            "display_order": safe_int(row.get("display_order")),
            "plo_map": {},
            "prereq_codes": [],
            "prev_codes": [],
        }

    # Đọc mapping course_plo.
    if not course_plo_df.empty:
        # Sắp theo sequence nếu có để kết quả ổn định.
        if "sequence_in_cell" in course_plo_df.columns:
            course_plo_df["_sequence"] = course_plo_df["sequence_in_cell"].map(lambda x: safe_int(x) if safe_int(x) is not None else 9999)
            course_plo_df = course_plo_df.sort_values(["course_id", "plo_id", "_sequence"])

        for _, row in course_plo_df.iterrows():
            if "is_mapped" in course_plo_df.columns and not (yes_no(row.get("is_mapped")) or safe_str(row.get("is_mapped")) == "1"):
                continue

            code = normalize_code(row.get("course_id"))
            plo = safe_str(row.get("plo_id")).upper()
            level = safe_str(row.get("level_code")).upper()

            if code not in courses or plo not in selected_plo_set or level not in LEVEL_ORDER:
                continue

            assess = yes_no(row.get("assessment")) or safe_str(row.get("assessment")) == "1"
            raw = safe_str(row.get("raw_value")) or (level + (",A" if assess else ""))

            # Nếu có trùng, ưu tiên mức cao hơn I<R<M hoặc bản có assessment.
            old = courses[code]["plo_map"].get(plo)
            new_info = {"level": level, "assess": assess, "raw": raw}
            if old is None:
                courses[code]["plo_map"][plo] = new_info
            else:
                old_score = LEVEL_ORDER.get(old["level"], 0) + (10 if old.get("assess") else 0)
                new_score = LEVEL_ORDER.get(level, 0) + (10 if assess else 0)
                if new_score > old_score:
                    courses[code]["plo_map"][plo] = new_info

    # Đọc loại quan hệ.
    relation_type_map = {}
    for _, row in relation_df.iterrows():
        relation_type_id = safe_str(row.get("relation_type_id"))
        name = safe_str(row.get("relation_type_name")).lower()
        source_column = safe_str(row.get("source_column")).lower()
        if not relation_type_id:
            continue
        if "tiên quyết" in name or "phải đạt" in name or "tien quyet" in name or "phai dat" in name or "tiên quyết" in source_column:
            relation_type_map[relation_type_id] = "prereq"
        elif "học trước" in name or "hoc truoc" in name or "học trước" in source_column:
            relation_type_map[relation_type_id] = "prev"
        else:
            relation_type_map[relation_type_id] = "other"

    # Đọc quan hệ học phần.
    for _, row in course_relation_df.iterrows():
        if "is_active" in course_relation_df.columns and not (yes_no(row.get("is_active")) or safe_str(row.get("is_active")) == "1"):
            continue

        target = normalize_code(row.get("target_course_id"))
        source = normalize_code(row.get("source_course_id"))
        relation_type_id = safe_str(row.get("relation_type_id"))

        if not target or not source or target not in courses:
            continue

        if source not in courses:
            warnings.append(f"Mã điều kiện {source} của {target} không có trong sheet course.")
            continue

        kind = relation_type_map.get(relation_type_id, "other")
        if kind == "prereq":
            if source not in courses[target]["prereq_codes"]:
                courses[target]["prereq_codes"].append(source)
        elif kind == "prev":
            if source not in courses[target]["prev_codes"]:
                courses[target]["prev_codes"].append(source)

    if yes_no(config.get("ONLY_PLO_RELATED_COURSES", "NO")):
        # Vẫn giữ các môn điều kiện của môn có PLO để click không bị mất dữ liệu.
        needed = set()
        for code, course in courses.items():
            if course["plo_map"]:
                needed.add(code)
                needed.update(course["prereq_codes"])
                needed.update(course["prev_codes"])
        courses = {k: v for k, v in courses.items() if k in needed}

    return courses, selected_plos, "database normalized", warnings

def group_courses_by_semester(courses):
    grouped = defaultdict(list)
    for c in courses.values():
        sem = c["semester"] if c["semester"] is not None else 0
        grouped[sem].append(c)
    for sem in grouped:
        grouped[sem].sort(key=lambda c: (c.get("display_order") if c.get("display_order") is not None else 9999, c["name"], c["code"]))
    return dict(sorted(grouped.items(), key=lambda x: x[0]))


def make_positions(grouped):
    positions = {}
    margin_x = 80
    margin_y = 118
    col_width = 310
    card_width = 250
    card_height = 92
    row_gap = 36
    semesters = list(grouped.keys())
    for col_index, semester in enumerate(semesters):
        x = margin_x + col_index * col_width
        for row_index, course in enumerate(grouped[semester]):
            y = margin_y + row_index * (card_height + row_gap)
            positions[course["code"]] = {"x": x, "y": y, "w": card_width, "h": card_height, "semester": semester}
    max_rows = max((len(v) for v in grouped.values()), default=1)
    width = margin_x * 2 + len(semesters) * col_width
    height = margin_y + max_rows * (card_height + row_gap) + 160
    return positions, width, height


def wrap_text(text, max_chars=32, max_lines=2):
    words = safe_str(text).split()
    lines, current = [], ""
    for word in words:
        if len((current + " " + word).strip()) <= max_chars:
            current = (current + " " + word).strip()
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines[:max_lines]


def course_tooltip(course):
    lines = [f"{course['code']} - {course['name']}"]
    lines.append(f"Học kỳ: {course['semester'] if course['semester'] is not None else 'Chưa rõ'}")
    if course["plo_map"]:
        lines.append("")
        lines.append("Đóng góp PLO:")
        for plo, info in sorted(course["plo_map"].items(), key=lambda x: x[0]):
            text = f"- {plo}: {info['level']} ({LEVEL_LABEL.get(info['level'], '')})"
            if info["assess"]:
                text += "; có đánh giá A"
            lines.append(text)
    else:
        lines.append("")
        lines.append("Không có mapping PLO trong các cột PLO đã chọn.")
    if course.get("prereq_codes") or course.get("prev_codes"):
        lines.append("")
        lines.append("Điều kiện của môn học này:")
        if course.get("prereq_codes"):
            lines.append("- Tiên quyết / phải đạt trước: " + ", ".join(course["prereq_codes"]))
        if course.get("prev_codes"):
            lines.append("- Học trước: " + ", ".join(course["prev_codes"]))
    return "\n".join(lines)


def badge_svg(course, selected_plos, x, y):
    """
    Nhãn PLO nhỏ chỉ dùng ở chế độ Hiện tất cả.
    Khi bấm riêng một PLO, các nhãn này sẽ được ẩn bằng JS để sơ đồ gọn.
    """
    parts = []
    bx = x
    max_x = x + 226
    shown = 0

    for plo in selected_plos:
        if plo not in course["plo_map"]:
            continue

        info = course["plo_map"][plo]
        level = info["level"]
        color = LEVEL_BORDER[level]
        label = f"{plo}:{level}{'A' if info['assess'] else ''}"
        badge_width = 46 + max(0, len(label) - 6) * 5

        if bx + badge_width > max_x:
            break

        parts.append(
            f'<rect x="{bx}" y="{y}" width="{badge_width}" height="18" rx="8" '
            f'fill="{color}" opacity="0.12" stroke="{color}" stroke-width="1"/>'
        )
        parts.append(
            f'<text x="{bx + 6}" y="{y + 13}" class="badge-text" fill="{BADGE_TEXT_COLOR}">{esc(label)}</text>'
        )
        bx += badge_width + 5
        shown += 1

    remaining = len(course["plo_map"]) - shown
    if remaining > 0 and bx + 32 <= max_x:
        parts.append(
            f'<text x="{bx}" y="{y + 13}" class="badge-text" fill="#6b7280">+{remaining}</text>'
        )

    return "".join(parts)


def make_svg(courses, grouped, positions, width, height, selected_plos):
    semesters = list(grouped.keys())
    parts = []
    legend_original_y = height - 80
    parts.append(f'''<svg id="flowSvg" xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}" data-base-width="{width}" data-base-height="{height}" data-margin-x="80" data-margin-y="118" data-col-width="310" data-card-width="250" data-card-height="92" data-row-gap="36" data-legend-y="{legend_original_y}">
<defs>
    <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
        <feDropShadow dx="0" dy="2" stdDeviation="2" flood-color="#000000" flood-opacity="0.12"/>
    </filter>
    <marker id="arrow-prereq" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto">
        <path d="M0,0 L0,6 L9,3 z" fill="{PREREQ_BORDER}" />
    </marker>
    <marker id="arrow-prev" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto">
        <path d="M0,0 L0,6 L9,3 z" fill="{PREV_BORDER}" />
    </marker>
    <marker id="arrow-both" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto">
        <path d="M0,0 L0,6 L9,3 z" fill="{BOTH_BORDER}" />
    </marker>
</defs>
<style>
    .title {{ font-family: Arial, sans-serif; font-size: 20px; font-weight: bold; fill: #111827; }}
    .subtitle {{ font-family: Arial, sans-serif; font-size: 13px; fill: #4b5563; }}
    .semester-title {{ font-family: Arial, sans-serif; font-size: 18px; font-weight: bold; fill: #1f2937; }}
    .course-card {{ filter: url(#shadow); cursor: pointer; transition: opacity 0.15s ease; }}
    .course-bg {{ stroke: {NORMAL_BORDER}; stroke-width: 2; transition: stroke 0.15s ease, stroke-width 0.15s ease, fill 0.15s ease; }}
    .assess-border {{ stroke: {ASSESS_COLOR}; stroke-width: 2.2; stroke-dasharray: 5 3; fill: none; opacity: 0; transition: opacity 0.15s ease; }}
    .relation-edge {{ fill: none; stroke-width: 3; opacity: 0.95; pointer-events: none; }}
    .course-code {{ font-family: Arial, sans-serif; font-size: 12px; font-weight: bold; fill: #374151; }}
    .course-name {{ font-family: Arial, sans-serif; font-size: 11px; fill: #111827; }}
    .badge-text {{ font-family: Arial, sans-serif; font-size: 10px; font-weight: bold; }}
    .plo-badges {{ transition: opacity 0.15s ease; }}
    .course-card.dim {{ opacity: 0.12; }}
    .course-card.unmapped {{ opacity: 0.08; }}
    .course-card.active {{ opacity: 1; }}
    .legend-text {{ font-family: Arial, sans-serif; font-size: 13px; fill: #374151; }}
</style>
<rect x="0" y="0" width="{width}" height="{height}" fill="#f8fafc" />
<text x="80" y="34" class="title">Sơ đồ tiến trình đào tạo theo PLO - database chuẩn hóa</text>
<text x="80" y="56" class="subtitle">Bấm nút PLO phía trên để hiện các học phần đóng góp cho PLO đó và các môn tiên quyết/học trước trực tiếp. Màu nền/viền: I = xanh lá, R = vàng, M = đỏ. Viền chấm = có đánh giá A.</text>
''')

    for col_index, semester in enumerate(semesters):
        x = 80 + col_index * 310
        title = "Chưa rõ HK" if semester == 0 else f"Học kỳ {semester}"
        parts.append(f'<g class="semester-column" data-semester="{semester}" data-x="{x}"><text x="{x}" y="98" class="semester-title">{esc(title)}</text><line x1="{x}" y1="108" x2="{x + 250}" y2="108" stroke="#cbd5e1" stroke-width="2"/></g>')

    parts.append('<g id="relationEdges"></g>')

    for semester, items in grouped.items():
        for course in items:
            code = course["code"]
            pos = positions[code]
            tooltip = course_tooltip(course)
            name_lines = wrap_text(course["name"])
            text_lines = [f'<text x="{pos["x"] + 12}" y="{pos["y"] + 21}" class="course-code">{esc(code)}</text>']
            for i, line in enumerate(name_lines):
                text_lines.append(f'<text x="{pos["x"] + 12}" y="{pos["y"] + 40 + i * 14}" class="course-name">{esc(line)}</text>')
            data_levels = {plo: info["level"] for plo, info in course["plo_map"].items()}
            data_assess = {plo: bool(info["assess"]) for plo, info in course["plo_map"].items()}
            # Ghi dạng chuỗi đơn giản để JS đọc nhanh: PLO1:I|PLO2:R
            levels_attr = "|".join([f"{plo}:{info['level']}" for plo, info in course["plo_map"].items()])
            assess_attr = "|".join([f"{plo}:1" for plo, info in course["plo_map"].items() if info["assess"]])
            prereq_attr = "|".join(course.get("prereq_codes", []))
            prev_attr = "|".join(course.get("prev_codes", []))
            parts.append(f'''<g id="course-{esc(code)}" class="course-card" data-code="{esc(code)}" data-semester="{pos["semester"]}" data-x="{pos["x"]}" data-y="{pos["y"]}" data-levels="{esc(levels_attr)}" data-assess="{esc(assess_attr)}" data-prereq="{esc(prereq_attr)}" data-prev="{esc(prev_attr)}" data-tooltip="{esc(tooltip)}">
                <rect class="course-bg" x="{pos["x"]}" y="{pos["y"]}" width="{pos["w"]}" height="{pos["h"]}" rx="10" ry="10" fill="{NORMAL_BG}"/>
                <rect class="assess-border" x="{pos["x"] + 4}" y="{pos["y"] + 4}" width="{pos["w"] - 8}" height="{pos["h"] - 8}" rx="9" ry="9"/>
                {''.join(text_lines)}
                <g class="plo-badges">{badge_svg(course, selected_plos, pos["x"] + 12, pos["y"] + 66)}</g>
            </g>''')

    legend_x = 80
    legend_y = height - 80
    parts.append(f'''<g id="legend">
        <rect x="{legend_x}" y="{legend_y}" width="26" height="18" rx="6" fill="{LEVEL_BG['I']}" stroke="{LEVEL_BORDER['I']}" stroke-width="3"/>
        <text x="{legend_x + 36}" y="{legend_y + 13}" class="legend-text">I - Introduction</text>
        <rect x="{legend_x + 190}" y="{legend_y}" width="26" height="18" rx="6" fill="{LEVEL_BG['R']}" stroke="{LEVEL_BORDER['R']}" stroke-width="3"/>
        <text x="{legend_x + 226}" y="{legend_y + 13}" class="legend-text">R - Reinforce</text>
        <rect x="{legend_x + 380}" y="{legend_y}" width="26" height="18" rx="6" fill="{LEVEL_BG['M']}" stroke="{LEVEL_BORDER['M']}" stroke-width="3"/>
        <text x="{legend_x + 416}" y="{legend_y + 13}" class="legend-text">M - Mastery</text>
        <rect x="{legend_x + 560}" y="{legend_y - 4}" width="38" height="26" rx="7" fill="white" stroke="{ASSESS_COLOR}" stroke-width="2.2" stroke-dasharray="5 3"/>
        <text x="{legend_x + 610}" y="{legend_y + 13}" class="legend-text">A - có đánh giá mức đạt PLO</text>
        <rect x="{legend_x}" y="{legend_y + 34}" width="26" height="18" rx="6" fill="{PREREQ_BG}" stroke="{PREREQ_BORDER}" stroke-width="3"/>
        <text x="{legend_x + 36}" y="{legend_y + 47}" class="legend-text">Môn tiên quyết / phải đạt trước được hiện kèm</text>
        <rect x="{legend_x + 340}" y="{legend_y + 34}" width="26" height="18" rx="6" fill="{PREV_BG}" stroke="{PREV_BORDER}" stroke-width="3"/>
        <text x="{legend_x + 376}" y="{legend_y + 47}" class="legend-text">Môn học trước được hiện kèm</text>
    </g></svg>''')
    return "\n".join(parts)


def make_html(svg_content, selected_plos):
    plo_buttons = "\n".join([f'<button class="plo-btn" data-plo="{esc(plo)}">{esc(plo)}</button>' for plo in selected_plos])
    return f'''<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<title>Sơ đồ tiến trình đào tạo theo PLO - database chuẩn hóa</title>
<style>
    body {{ margin: 0; font-family: Arial, sans-serif; background: #f8fafc; color: #111827; }}
    .topbar {{ position: sticky; top: 0; z-index: 10; background: white; border-bottom: 1px solid #e5e7eb; padding: 12px 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); }}
    .topbar h1 {{ margin: 0 0 8px 0; font-size: 20px; }}
    .toolbar {{ display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }}
    .plo-btn, .reset-btn {{ border: 1px solid #cbd5e1; background: #ffffff; border-radius: 999px; padding: 7px 13px; font-weight: 700; cursor: pointer; color: #1f2937; }}
    .plo-btn:hover, .reset-btn:hover {{ background: #f1f5f9; }}
    .plo-btn.active {{ background: #111827; color: #ffffff; border-color: #111827; }}
    .status {{ margin-left: 8px; font-size: 13px; color: #4b5563; }}
    .canvas {{ overflow: auto; height: calc(100vh - 104px); padding: 16px; }}
    .cm-hover-tooltip {{ position: fixed; display: none; max-width: 660px; max-height: 70vh; overflow: auto; background: #111827; color: white; padding: 10px 12px; border-radius: 8px; font-size: 13px; line-height: 1.45; z-index: 9999; pointer-events: none; white-space: pre-line; box-shadow: 0 8px 24px rgba(0,0,0,0.18); }}
</style>
</head>
<body>
<div class="topbar">
    <h1>Sơ đồ tiến trình đào tạo theo PLO - database chuẩn hóa</h1>
    <div class="toolbar">
        {plo_buttons}
        <button class="reset-btn" id="resetBtn">Hiện tất cả</button>
        <span class="status" id="statusText">Chọn một PLO để lọc học phần.</span>
    </div>
</div>
<div class="canvas">{svg_content}</div>
<div id="cm-hover-tooltip" class="cm-hover-tooltip"></div>
<script>
const svg = document.getElementById('flowSvg');
const cards = Array.from(document.querySelectorAll('.course-card'));
const columns = Array.from(document.querySelectorAll('.semester-column'));
const buttons = document.querySelectorAll('.plo-btn');
const resetBtn = document.getElementById('resetBtn');
const statusText = document.getElementById('statusText');
const tooltip = document.getElementById('cm-hover-tooltip');
const legend = document.getElementById('legend');
const relationEdges = document.getElementById('relationEdges');
const LEVEL_BORDER = {{ I: '#22c55e', R: '#facc15', M: '#dc2626' }};
const LEVEL_BG = {{ I: '#dcfce7', R: '#fef3c7', M: '#fecaca' }};
const NORMAL_BORDER = '{NORMAL_BORDER}';
const NORMAL_BG = '{NORMAL_BG}';
const PREREQ_BORDER = '{PREREQ_BORDER}';
const PREREQ_BG = '{PREREQ_BG}';
const PREV_BORDER = '{PREV_BORDER}';
const PREV_BG = '{PREV_BG}';
const BOTH_BORDER = '{BOTH_BORDER}';
const BOTH_BG = '{BOTH_BG}';
let selectedPlo = null;
let selectedCourseCode = null;

const layout = {{
  baseWidth: Number(svg.dataset.baseWidth),
  baseHeight: Number(svg.dataset.baseHeight),
  marginX: Number(svg.dataset.marginX),
  marginY: Number(svg.dataset.marginY),
  colWidth: Number(svg.dataset.colWidth),
  cardWidth: Number(svg.dataset.cardWidth),
  cardHeight: Number(svg.dataset.cardHeight),
  rowGap: Number(svg.dataset.rowGap),
  legendY: Number(svg.dataset.legendY)
}};

function parsePairs(text) {{
  const obj = {{}};
  if (!text) return obj;
  text.split('|').forEach(part => {{
    const pieces = part.split(':');
    if (pieces.length === 2) obj[pieces[0]] = pieces[1];
  }});
  return obj;
}}

function parseList(text) {{
  if (!text) return [];
  return text.split('|').map(x => x.trim()).filter(Boolean);
}}

const cardByCode = new Map(cards.map(card => [card.dataset.code, card]));

function clearRelationEdges() {{
  if (relationEdges) relationEdges.innerHTML = '';
}}

function resetCardStyle(card) {{
  const bg = card.querySelector('.course-bg');
  const assess = card.querySelector('.assess-border');
  const badges = card.querySelector('.plo-badges');
  card.classList.remove('dim', 'unmapped', 'active', 'selected-course');
  card.style.display = '';
  card.setAttribute('transform', '');
  bg.style.stroke = NORMAL_BORDER;
  bg.style.strokeWidth = '2';
  bg.style.fill = NORMAL_BG;
  assess.style.opacity = 0;
  // Khi quay về “Hiện tất cả”, phải bật lại toàn bộ nhãn PLOx:I/R/M.
  // Trước đó applyPlo() đã đặt badges.style.display = 'none' cho PLO được lọc.
  if (badges) badges.style.display = '';
}}

function setSvgSize(width, height) {{
  svg.setAttribute('width', width);
  svg.setAttribute('height', height);
  svg.setAttribute('viewBox', `0 0 ${{width}} ${{height}}`);
}}

function moveLegend(newHeight) {{
  if (!legend) return;
  const newY = Math.max(layout.marginY + 40, newHeight - 80);
  legend.setAttribute('transform', `translate(0, ${{newY - layout.legendY}})`);
}}

function restoreFullLayout() {{
  setSvgSize(layout.baseWidth, layout.baseHeight);
  if (legend) legend.setAttribute('transform', '');
  columns.forEach(col => {{
    col.style.display = '';
    col.setAttribute('transform', '');
  }});
  cards.forEach(card => {{
    card.dataset.currentX = card.dataset.x;
    card.dataset.currentY = card.dataset.y;
  }});
  clearRelationEdges();
}}

function applyCompactLayout(visibleCards) {{
  const visibleBySemester = new Map();
  visibleCards.forEach(card => {{
    const sem = card.dataset.semester || '0';
    if (!visibleBySemester.has(sem)) visibleBySemester.set(sem, []);
    visibleBySemester.get(sem).push(card);
  }});

  const visibleSemesters = Array.from(visibleBySemester.keys()).sort((a, b) => Number(a) - Number(b));

  columns.forEach(col => {{
    const sem = col.dataset.semester || '0';
    const colIndex = visibleSemesters.indexOf(sem);
    if (colIndex === -1) {{
      col.style.display = 'none';
      return;
    }}
    const originalX = Number(col.dataset.x);
    const newX = layout.marginX + colIndex * layout.colWidth;
    col.style.display = '';
    col.setAttribute('transform', `translate(${{newX - originalX}}, 0)`);
  }});

  let maxRows = 1;
  visibleSemesters.forEach((sem, colIndex) => {{
    const list = visibleBySemester.get(sem).sort((a, b) => Number(a.dataset.y) - Number(b.dataset.y));
    maxRows = Math.max(maxRows, list.length);
    list.forEach((card, rowIndex) => {{
      const originalX = Number(card.dataset.x);
      const originalY = Number(card.dataset.y);
      const newX = layout.marginX + colIndex * layout.colWidth;
      const newY = layout.marginY + rowIndex * (layout.cardHeight + layout.rowGap);
      card.dataset.currentX = String(newX);
      card.dataset.currentY = String(newY);
      card.setAttribute('transform', `translate(${{newX - originalX}}, ${{newY - originalY}})`);
    }});
  }});

  const newWidth = layout.marginX * 2 + visibleSemesters.length * layout.colWidth;
  const newHeight = layout.marginY + maxRows * (layout.cardHeight + layout.rowGap) + 160;
  setSvgSize(Math.max(newWidth, 900), newHeight);
  moveLegend(newHeight);
}}

function styleDirectCard(card, plo, levelCounts) {{
  const levels = parsePairs(card.dataset.levels);
  const assess = parsePairs(card.dataset.assess);
  const bg = card.querySelector('.course-bg');
  const assessBorder = card.querySelector('.assess-border');
  const badges = card.querySelector('.plo-badges');
  const level = levels[plo];

  card.style.display = '';
  card.classList.add('active');
  if (badges) badges.style.display = 'none';

  if (level) {{
    if (levelCounts) levelCounts[level] += 1;
    bg.style.stroke = LEVEL_BORDER[level];
    bg.style.strokeWidth = '3.5';
    bg.style.fill = LEVEL_BG[level];
    if (assess[plo]) assessBorder.style.opacity = 1;
  }}
}}

function styleRelationOnlyCard(card, relationKind) {{
  const bg = card.querySelector('.course-bg');
  const badges = card.querySelector('.plo-badges');

  card.style.display = '';
  card.classList.add('active');
  if (badges) badges.style.display = 'none';

  // Môn điều kiện chỉ hiển thị bằng ô bình thường; mối liên hệ được thể hiện bằng mũi tên.
  bg.style.stroke = NORMAL_BORDER;
  bg.style.fill = NORMAL_BG;
  bg.style.strokeWidth = '2';
}}

function cardPosition(card) {{
  return {{
    x: Number(card.dataset.currentX || card.dataset.x),
    y: Number(card.dataset.currentY || card.dataset.y),
    w: layout.cardWidth,
    h: layout.cardHeight
  }};
}}

function edgePath(sourceCard, targetCard) {{
  const source = cardPosition(sourceCard);
  const target = cardPosition(targetCard);
  const sx = source.x + source.w;
  const sy = source.y + source.h / 2;
  const tx = target.x;
  const ty = target.y + target.h / 2;

  if (tx <= sx) {{
    const offset = 55;
    const p1x = sx + offset;
    const p2x = tx - offset;
    const midY = ty >= sy ? Math.max(sy, ty) + 44 : Math.min(sy, ty) - 44;
    return `M ${{sx}} ${{sy}} C ${{p1x}} ${{sy}}, ${{p1x}} ${{midY}}, ${{(sx + tx) / 2}} ${{midY}} S ${{p2x}} ${{ty}}, ${{tx}} ${{ty}}`;
  }}

  const dx = Math.max(80, (tx - sx) / 2);
  return `M ${{sx}} ${{sy}} C ${{sx + dx}} ${{sy}}, ${{tx - dx}} ${{ty}}, ${{tx}} ${{ty}}`;
}}

function drawRelationArrow(sourceCard, targetCard, relationKind) {{
  if (!relationEdges || !sourceCard || !targetCard) return;
  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  const color = relationKind === 'both' ? BOTH_BORDER : (relationKind === 'prereq' ? PREREQ_BORDER : PREV_BORDER);
  const marker = relationKind === 'both' ? 'arrow-both' : (relationKind === 'prereq' ? 'arrow-prereq' : 'arrow-prev');
  path.setAttribute('class', 'relation-edge');
  path.setAttribute('d', edgePath(sourceCard, targetCard));
  path.setAttribute('stroke', color);
  path.setAttribute('marker-end', `url(#${{marker}})`);
  if (relationKind === 'prev') path.setAttribute('stroke-dasharray', '7 4');
  relationEdges.appendChild(path);
}}

function directCardsForPlo(plo) {{
  return cards.filter(card => {{
    const levels = parsePairs(card.dataset.levels);
    return Boolean(levels[plo]);
  }});
}}

function applyPlo(plo) {{
  selectedPlo = plo;
  selectedCourseCode = null;
  clearRelationEdges();
  let count = 0;
  const visibleCards = [];
  const levelCounts = {{ I: 0, R: 0, M: 0 }};
  buttons.forEach(btn => btn.classList.toggle('active', btn.dataset.plo === plo));

  cards.forEach(card => {{
    resetCardStyle(card);
    const levels = parsePairs(card.dataset.levels);
    const level = levels[plo];

    if (!level) {{
      card.style.display = 'none';
      return;
    }}

    count += 1;
    visibleCards.push(card);
    styleDirectCard(card, plo, levelCounts);
  }});

  applyCompactLayout(visibleCards);
  statusText.textContent = `${{plo}}: đang hiển thị ${{count}} học phần đóng góp trực tiếp — I: ${{levelCounts.I}}, R: ${{levelCounts.R}}, M: ${{levelCounts.M}}. Click vào một môn để hiện mũi tên từ môn tiên quyết/học trước sang môn đó.`;
}}

function showPrereqForCourse(card) {{
  if (!selectedPlo) return;
  clearRelationEdges();

  const levels = parsePairs(card.dataset.levels);
  if (!levels[selectedPlo]) return;

  selectedCourseCode = card.dataset.code;

  const directCards = directCardsForPlo(selectedPlo);
  const visibleCards = [];
  const visibleSet = new Set();
  const levelCounts = {{ I: 0, R: 0, M: 0 }};

  const prereqCodes = new Set(parseList(card.dataset.prereq).filter(code => cardByCode.has(code)));
  const prevCodes = new Set(parseList(card.dataset.prev).filter(code => cardByCode.has(code)));

  cards.forEach(resetCardStyle);

  directCards.forEach(directCard => {{
    visibleCards.push(directCard);
    visibleSet.add(directCard.dataset.code);
    styleDirectCard(directCard, selectedPlo, levelCounts);
  }});

  let prereqCount = 0;
  let prevCount = 0;

  const relationCodes = new Set([...prereqCodes, ...prevCodes]);
  relationCodes.forEach(code => {{
    const relCard = cardByCode.get(code);
    if (!relCard) return;

    const isPrereq = prereqCodes.has(code);
    const isPrev = prevCodes.has(code);
    const isAlreadyDirect = visibleSet.has(code);

    if (!isAlreadyDirect) {{
      visibleCards.push(relCard);
      visibleSet.add(code);
      if (isPrereq && isPrev) {{
        styleRelationOnlyCard(relCard, 'both');
      }} else if (isPrereq) {{
        styleRelationOnlyCard(relCard, 'prereq');
      }} else {{
        styleRelationOnlyCard(relCard, 'prev');
      }}
    }}

    if (isPrereq) prereqCount += 1;
    if (isPrev) prevCount += 1;
  }});

  cards.forEach(c => {{
    if (!visibleSet.has(c.dataset.code)) c.style.display = 'none';
  }});

  applyCompactLayout(visibleCards);

  relationCodes.forEach(code => {{
    const relCard = cardByCode.get(code);
    if (!relCard || !visibleSet.has(code)) return;
    const isPrereq = prereqCodes.has(code);
    const isPrev = prevCodes.has(code);
    let kind = 'prev';
    if (isPrereq && isPrev) kind = 'both';
    else if (isPrereq) kind = 'prereq';
    drawRelationArrow(relCard, card, kind);
  }});

  statusText.textContent = `${{selectedPlo}}: đang hiển thị các môn đóng góp trực tiếp. Đã chọn ${{card.dataset.code}}; mũi tên cho biết ${{prereqCount}} môn tiên quyết/phải đạt trước và ${{prevCount}} môn học trước của môn này.`;
}}

function showAll() {{
  selectedPlo = null;
  selectedCourseCode = null;
  clearRelationEdges();
  buttons.forEach(btn => btn.classList.remove('active'));
  cards.forEach(resetCardStyle);
  restoreFullLayout();
  statusText.textContent = 'Đang hiển thị tất cả học phần. Chọn một PLO để chỉ hiện các học phần liên quan.';
}}

function showTooltip(card, event) {{
  const text = card.dataset.tooltip || '';
  if (!text) return;
  tooltip.textContent = text;
  tooltip.style.display = 'block';
  moveTooltip(event);
}}
function hideTooltip() {{ tooltip.style.display = 'none'; }}
function moveTooltip(event) {{
  const padding = 16, offset = 14;
  tooltip.style.left = event.clientX + offset + 'px';
  tooltip.style.top = event.clientY + offset + 'px';
  const rect = tooltip.getBoundingClientRect();
  if (rect.right > window.innerWidth - padding) tooltip.style.left = event.clientX - rect.width - offset + 'px';
  if (rect.bottom > window.innerHeight - padding) tooltip.style.top = event.clientY - rect.height - offset + 'px';
}}

buttons.forEach(btn => btn.addEventListener('click', () => applyPlo(btn.dataset.plo)));
resetBtn.addEventListener('click', showAll);
cards.forEach(card => {{
  card.addEventListener('mouseenter', e => showTooltip(card, e));
  card.addEventListener('mousemove', e => {{ if (tooltip.style.display === 'block') moveTooltip(e); }});
  card.addEventListener('mouseleave', hideTooltip);
  card.addEventListener('click', e => {{
    e.stopPropagation();
    showPrereqForCourse(card);
  }});
}});
document.addEventListener('keydown', e => {{ if (e.key === 'Escape') showAll(); }});
</script>
</body>
</html>'''


def main():
    input_file = Path(sys.argv[1]) if len(sys.argv) >= 2 else Path(INPUT_DEFAULT)
    cli_plos = None
    if len(sys.argv) >= 3:
        cli_plos = [x.strip() for x in re.split(r"[,;\n]+", sys.argv[2]) if x.strip()]
    if not input_file.exists():
        print(f"Không tìm thấy file: {input_file}")
        print('Cách dùng: python draw_flow_plo_database.py "progress.xlsx" "PLO1,PLO2"')
        sys.exit(1)

    config = read_config(input_file)
    courses, selected_plos, sheet, warnings = read_courses(input_file, config, cli_plos)
    grouped = group_courses_by_semester(courses)
    positions, width, height = make_positions(grouped)
    svg = make_svg(courses, grouped, positions, width, height, selected_plos)
    html_content = make_html(svg, selected_plos)

    output_svg = input_file.with_name(OUTPUT_SVG)
    output_html = input_file.with_name(OUTPUT_HTML)
    output_svg.write_text(svg, encoding="utf-8")
    output_html.write_text(html_content, encoding="utf-8")

    mapped_courses = sum(1 for c in courses.values() if c["plo_map"])
    print("Đã tạo sơ đồ PLO từ file database chuẩn hóa: click vào môn sẽ hiện mũi tên từ môn tiên quyết/học trước sang môn được chọn.")
    print(f"Sheet dữ liệu: {sheet}")
    print(f"PLO: {', '.join(selected_plos)}")
    print(f"HTML: {output_html}")
    print(f"SVG : {output_svg}")
    print(f"Số học phần: {len(courses)}")
    print(f"Số học phần có mapping PLO: {mapped_courses}")
    if warnings:
        print("Cảnh báo:")
        for w in warnings[:10]:
            print("- " + w)


if __name__ == "__main__":
    main()
