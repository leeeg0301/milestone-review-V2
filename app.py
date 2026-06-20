import io
import re
import itertools
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from matplotlib import font_manager


# =========================================================
# 1. 기본 설정
# =========================================================
ROAD_START = 0.0
ROAD_END = 106.84

IC_POINTS = {
    "서영암IC": 0,
    "강진IC": 20,
    "장흥IC": 40,
    "보성IC": 60,
    "벌교IC": 80,
    "남순천": 100,
    "해룡IC": 106.84,
}

LANES = ["1차로", "2차로", "갓길"]


# =========================================================
# 2. 한글 폰트 설정
# =========================================================
def set_korean_font():
    available_fonts = {f.name for f in font_manager.fontManager.ttflist}

    font_candidates = [
        "Malgun Gothic",
        "AppleGothic",
        "NanumGothic",
        "Noto Sans CJK KR",
        "Noto Sans KR",
        "DejaVu Sans",
    ]

    for font_name in font_candidates:
        if font_name in available_fonts:
            plt.rcParams["font.family"] = font_name
            break

    plt.rcParams["axes.unicode_minus"] = False


set_korean_font()


# =========================================================
# 3. 기본 문자열 처리 함수
# =========================================================
def clean_text(value):
    """
    NaN, None, null 등을 빈 문자열로 정리.
    """
    if pd.isna(value):
        return ""

    text = str(value).strip()

    if text.lower() in ["nan", "none", "null"]:
        return ""

    return text


def clean_group_name(value):
    """
    그룹명 빈칸은 그룹 없음으로 처리.
    빈칸끼리 nan 그룹으로 묶이는 문제 방지.
    """
    text = clean_text(value)

    if text == "":
        return ""

    return text


def make_unique_columns(columns):
    """
    엑셀 헤더 중 빈 컬럼명 또는 중복 컬럼명을 안전하게 정리.
    """
    result = []
    seen = {}

    for i, col in enumerate(columns):
        name = clean_text(col)

        if name == "":
            name = f"빈컬럼_{i + 1}"

        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0

        result.append(name)

    return result


# =========================================================
# 4. 엑셀 읽기: 헤더 자동 탐색
# =========================================================
def read_excel_smart(uploaded_file, sheet_name):
    """
    도로공사 작업계획서처럼 상단에 제목/공백이 있고,
    실제 헤더가 중간에 있는 엑셀을 자동으로 읽음.

    실제 헤더 행에서 아래 항목을 탐색:
    - 공사명
    - 방향
    - 공사구간
    - 차단차로
    """
    uploaded_file.seek(0)

    probe_df = pd.read_excel(
        uploaded_file,
        sheet_name=sheet_name,
        header=None,
    )

    header_row_idx = None

    for idx, row in probe_df.iterrows():
        values = [clean_text(v).replace(" ", "") for v in row.tolist()]

        has_work_name = any("공사명" in v for v in values)
        has_direction = any("방향" in v for v in values)
        has_section = any("공사구간" in v or "구간" in v or "이정" in v for v in values)
        has_lane = any("차단차로" in v or "차로" in v for v in values)

        score = sum([has_work_name, has_direction, has_section, has_lane])

        if score >= 3:
            header_row_idx = idx
            break

    if header_row_idx is None:
        return pd.DataFrame()

    columns = make_unique_columns(probe_df.iloc[header_row_idx].tolist())

    raw_df = probe_df.iloc[header_row_idx + 1:].copy()
    raw_df.columns = columns
    raw_df = raw_df.dropna(how="all").reset_index(drop=True)

    return raw_df


# =========================================================
# 5. 컬럼 자동 추정
# =========================================================
def guess_column(columns, keywords):
    """
    컬럼명 목록에서 원하는 키워드가 들어간 컬럼을 자동 선택.
    """
    for col in columns:
        col_text = clean_text(col).replace(" ", "")

        for keyword in keywords:
            if keyword in col_text:
                return col

    return columns[0] if len(columns) > 0 else None


# =========================================================
# 6. 방향 / 이정 / 차로 파싱 함수
# =========================================================
def parse_direction(value):
    """
    엑셀 방향값을 내부 방향값으로 변환.

    예:
    - 순천종점 -> ["순천"]
    - 영암기점 -> ["영암"]
    - 양방향 -> ["순천", "영암"]
    """
    text = clean_text(value).replace(" ", "")

    if text == "":
        return []

    if "양방향" in text or text == "양":
        return ["순천", "영암"]

    if "순천" in text:
        return ["순천"]

    if "영암" in text:
        return ["영암"]

    return []


def parse_km_interval(value):
    """
    공사구간 문자열에서 이정 숫자 2개를 추출.

    예:
    - 103.18km ~ 104.87km
    - 49.3km ~ 62.7km
    - 95km ~ 2.3km
    - 0km ~ 106.84km

    반환:
    - (작은 이정, 큰 이정)
    """
    text = clean_text(value)

    if text == "":
        return None

    text = text.replace(",", "")

    # 1차: km/k/㎞ 앞 숫자를 우선 추출
    km_pattern = r"(\d+(?:\.\d+)?)\s*(?:k|K|km|KM|㎞)"
    nums = re.findall(km_pattern, text)

    # 2차: k 표기가 없으면 일반 숫자 추출
    if len(nums) < 2:
        nums = re.findall(r"\d+(?:\.\d+)?", text)

    if len(nums) < 2:
        return None

    start = float(nums[0])
    end = float(nums[1])

    draw_start = min(start, end)
    draw_end = max(start, end)

    # 관리구간과 아예 겹치지 않으면 제외
    if draw_end <= ROAD_START or draw_start >= ROAD_END:
        return None

    # 관리구간 밖으로 살짝 넘어가는 경우 잘라냄
    draw_start = max(draw_start, ROAD_START)
    draw_end = min(draw_end, ROAD_END)

    if draw_end <= draw_start:
        return None

    return draw_start, draw_end


def is_full_range(start, end):
    """
    0~106.84 전체 구간이면 기본 미표시 처리.
    약간의 오차 허용.
    """
    return start <= 0.1 and end >= ROAD_END - 0.2


def parse_lanes(value):
    """
    차단차로 컬럼에서 1차로, 2차로, 갓길만 추출.
    이동차단은 차로로 표시하지 않음.
    """
    text = clean_text(value).replace(" ", "")

    if text == "":
        return []

    lanes = []

    if "1차로" in text or "1차" in text:
        lanes.append("1차로")

    if "2차로" in text or "2차" in text:
        lanes.append("2차로")

    if "갓길" in text:
        lanes.append("갓길")

    # LANES 순서대로 중복 제거
    result = []
    for lane in LANES:
        if lane in lanes and lane not in result:
            result.append(lane)

    return result


def has_moving_closure(value):
    """
    이동차단 포함 여부 확인.
    이동차단은 광범위하게 잡히는 경우가 많아 기본 미표시 처리에 활용.
    """
    text = clean_text(value).replace(" ", "")

    return "이동차단" in text


def lane_text_to_list(value):
    """
    사용자가 data_editor에서 수정한 차로 문자열을 다시 리스트로 변환.
    """
    return parse_lanes(value)


# =========================================================
# 7. 엑셀 원본 -> 작업표 변환
# =========================================================
def parse_excel_to_work_table(
    raw_df,
    name_col,
    direction_col,
    section_col,
    lane_col,
    hide_full_range=True,
    hide_moving_closure=True,
    hide_no_lane=True,
):
    """
    엑셀 원본에서 도식에 필요한 데이터만 추출.

    양방향은 순천/영암 2개 행으로 분리.
    """
    rows = []
    auto_no = 1

    for raw_idx, row in raw_df.iterrows():
        work_name = clean_text(row.get(name_col, ""))
        direction_raw = clean_text(row.get(direction_col, ""))
        section_raw = clean_text(row.get(section_col, ""))
        lane_raw = clean_text(row.get(lane_col, ""))

        # 공사명도 없고 구간도 없으면 데이터 행이 아니라고 판단
        if work_name == "" and section_raw == "":
            continue

        interval = parse_km_interval(section_raw)

        if interval is None:
            continue

        start, end = interval

        directions = parse_direction(direction_raw)

        if not directions:
            continue

        lanes = parse_lanes(lane_raw)

        default_display = True
        hide_reason_list = []

        if hide_full_range and is_full_range(start, end):
            default_display = False
            hide_reason_list.append("전체구간")

        if hide_moving_closure and has_moving_closure(lane_raw):
            default_display = False
            hide_reason_list.append("이동차단")

        if hide_no_lane and len(lanes) == 0:
            default_display = False
            hide_reason_list.append("차로정보없음")

        hide_reason = ",".join(hide_reason_list)

        for direction in directions:
            rows.append({
                "표시여부": default_display,
                "제외사유": hide_reason,
                "번호": auto_no,
                "공사명": work_name,
                "방향": direction,
                "시점": start,
                "종점": end,
                "차로": ",".join(lanes),
                "그룹명": "",
                "원본행": raw_idx + 1,
                "원문방향": direction_raw,
                "원문공사구간": section_raw,
                "원문차단차로": lane_raw,
            })

            auto_no += 1

    return pd.DataFrame(rows)


# =========================================================
# 8. 작업표 -> 내부 계산용 표
# =========================================================
def normalize_work_table(df):
    """
    사용자가 수정한 작업표에서 표시여부=True인 행만 내부 계산용으로 정리.
    """
    rows = []

    if df.empty:
        return pd.DataFrame()

    for _, row in df.iterrows():
        show = bool(row.get("표시여부", True))

        if not show:
            continue

        if pd.isna(row.get("번호")) or pd.isna(row.get("시점")) or pd.isna(row.get("종점")):
            continue

        try:
            start = float(row["시점"])
            end = float(row["종점"])
        except ValueError:
            continue

        draw_start = min(start, end)
        draw_end = max(start, end)

        if draw_end <= ROAD_START or draw_start >= ROAD_END:
            continue

        draw_start = max(draw_start, ROAD_START)
        draw_end = min(draw_end, ROAD_END)

        if draw_end <= draw_start:
            continue

        direction = clean_text(row.get("방향", ""))

        if direction not in ["순천", "영암"]:
            continue

        lanes = lane_text_to_list(row.get("차로", ""))

        if len(lanes) == 0:
            continue

        group = clean_group_name(row.get("그룹명", ""))

        no = clean_text(row.get("번호", ""))
        no = no.replace("#", "")

        name = clean_text(row.get("공사명", ""))

        rows.append({
            "번호": no,
            "공사명": name,
            "방향": direction,
            "시점": draw_start,
            "종점": draw_end,
            "차로": lanes,
            "차로표시": ",".join(lanes),
            "그룹명": group,
        })

    return pd.DataFrame(rows)


# =========================================================
# 9. 다공종 그룹화
# =========================================================
def build_work_units(df, use_group=True):
    """
    개별 공사를 실제 검토 단위로 변환.

    그룹명이 같은 경우:
    - 같은 그룹명
    - 같은 방향
    을 하나의 다공종 작업으로 묶음.
    """
    if df.empty:
        return pd.DataFrame()

    units = {}

    for idx, row in df.iterrows():
        group_name = clean_group_name(row.get("그룹명", ""))

        if use_group and group_name != "":
            key = ("GROUP", group_name, row["방향"])
        else:
            key = ("SINGLE", idx, row["방향"])

        if key not in units:
            units[key] = {
                "unit_id": "|".join(map(str, key)),
                "번호목록": [],
                "공사명목록": [],
                "방향": row["방향"],
                "시점": row["시점"],
                "종점": row["종점"],
                "차로": set(),
                "그룹명": group_name,
            }

        units[key]["번호목록"].append(row["번호"])
        units[key]["공사명목록"].append(row["공사명"])
        units[key]["시점"] = min(units[key]["시점"], row["시점"])
        units[key]["종점"] = max(units[key]["종점"], row["종점"])
        units[key]["차로"].update(row["차로"])

    result = []

    for _, unit in units.items():
        lane_list = [lane for lane in LANES if lane in unit["차로"]]
        no_list = unit["번호목록"]
        name_list = [name for name in unit["공사명목록"] if name]

        is_group = len(no_list) >= 2 and unit["그룹명"] != ""

        if is_group:
            display_no = ",".join([f"#{n}" for n in no_list])
            display_name = "다공종작업"
            detail_name = " / ".join(name_list)
        else:
            display_no = f"#{no_list[0]}"
            display_name = name_list[0] if name_list else ""
            detail_name = display_name

        result.append({
            "unit_id": unit["unit_id"],
            "번호표시": display_no,
            "공사명": display_name,
            "상세공사명": detail_name,
            "방향": unit["방향"],
            "시점": unit["시점"],
            "종점": unit["종점"],
            "차로": lane_list,
            "차로표시": ",".join(lane_list),
            "그룹명": unit["그룹명"],
            "다공종여부": is_group,
        })

    return pd.DataFrame(result)


# =========================================================
# 10. 겹침 / 인접 판정
# =========================================================
def interval_relation(a_start, a_end, b_start, b_end):
    """
    두 구간의 겹침 또는 이격거리 계산.
    """
    overlap_start = max(a_start, b_start)
    overlap_end = min(a_end, b_end)
    overlap_len = overlap_end - overlap_start

    if overlap_len > 0:
        return {
            "type": "overlap",
            "distance": 0.0,
            "start": overlap_start,
            "end": overlap_end,
        }

    gap = max(a_start, b_start) - min(a_end, b_end)

    gap_start = min(a_end, b_end)
    gap_end = max(a_start, b_start)

    return {
        "type": "near",
        "distance": gap,
        "start": gap_start,
        "end": gap_end,
    }


def find_conflicts(units, threshold_km=5.0, same_direction_only=True, consider_lane=False):
    """
    모든 작업 단위를 2개씩 비교해서 겹침 또는 기준 km 이내 인접 여부를 찾음.
    """
    if units.empty:
        return pd.DataFrame()

    conflicts = []

    for i, j in itertools.combinations(units.index, 2):
        a = units.loc[i]
        b = units.loc[j]

        if same_direction_only and a["방향"] != b["방향"]:
            continue

        if consider_lane:
            common_lanes = set(a["차로"]) & set(b["차로"])

            if len(common_lanes) == 0:
                continue

        relation = interval_relation(a["시점"], a["종점"], b["시점"], b["종점"])

        if relation["type"] == "overlap":
            conflicts.append({
                "작업1": f"{a['번호표시']} {a['공사명']}",
                "작업2": f"{b['번호표시']} {b['공사명']}",
                "방향": a["방향"] if a["방향"] == b["방향"] else "양방향",
                "구분": "구간 겹침",
                "문제구간": f"{relation['start']:.1f}k ~ {relation['end']:.1f}k",
                "이격거리(km)": 0.0,
                "start": relation["start"],
                "end": relation["end"],
                "type": "overlap",
                "unit_id1": a["unit_id"],
                "unit_id2": b["unit_id"],
            })

        elif relation["distance"] <= threshold_km:
            conflicts.append({
                "작업1": f"{a['번호표시']} {a['공사명']}",
                "작업2": f"{b['번호표시']} {b['공사명']}",
                "방향": a["방향"] if a["방향"] == b["방향"] else "양방향",
                "구분": f"{threshold_km:g}km 이내 인접",
                "문제구간": f"{relation['start']:.1f}k ~ {relation['end']:.1f}k",
                "이격거리(km)": round(relation["distance"], 2),
                "start": relation["start"],
                "end": relation["end"],
                "type": "near",
                "unit_id1": a["unit_id"],
                "unit_id2": b["unit_id"],
            })

    return pd.DataFrame(conflicts)


# =========================================================
# 11. 도식 그리기
# =========================================================
def get_lane_y_range(direction, lanes):
    """
    방향과 차로를 y좌표로 변환.

    영암방향: 중앙선 위
    순천방향: 중앙선 아래
    """
    lane_idx = {
        "1차로": 0,
        "2차로": 1,
        "갓길": 2,
    }

    idxs = [lane_idx[lane] for lane in lanes if lane in lane_idx]

    if not idxs:
        return None

    min_idx = min(idxs)
    max_idx = max(idxs)

    if direction == "영암":
        y0 = min_idx
        y1 = max_idx + 1
    else:
        y0 = -(max_idx + 1)
        y1 = -min_idx

    return y0, y1


def draw_diagram(units, conflicts, show_warnings=True, submit_mode=False):
    """
    공사구간 도식 생성.
    """
    fig, ax = plt.subplots(figsize=(15, 4.8), dpi=160)

    ax.set_xlim(-2, 108.6)
    ax.set_ylim(-3.85, 4.05)
    ax.axis("off")

    # 검토용일 때만 경고 음영 표시
    if show_warnings and conflicts is not None and not conflicts.empty:
        for _, c in conflicts.iterrows():
            if c["type"] == "overlap":
                ax.axvspan(c["start"], c["end"], color="red", alpha=0.18, zorder=0)
            else:
                ax.axvspan(c["start"], c["end"], color="orange", alpha=0.13, zorder=0)

    # 세로 격자선
    x_ticks = list(range(0, 101, 10)) + [ROAD_END]

    for x in x_ticks:
        major = x in [0, 20, 40, 60, 80, 100] or abs(x - ROAD_END) < 0.01
        lw = 1.2 if major else 0.8
        ax.plot([x, x], [-3, 3], color="black", linewidth=lw, alpha=0.85)

    # 가로 차로선
    for y in [-3, -2, -1, 0, 1, 2, 3]:
        if y == 0:
            ax.plot([0, ROAD_END], [y, y], color="black", linewidth=3.2)
        else:
            ax.plot([0, ROAD_END], [y, y], color="black", linewidth=0.9, alpha=0.8)

    # 이정 숫자
    for x in range(0, 101, 10):
        label = "0k" if x == 0 else f"{x}"
        ax.text(x + 0.4, 3.08, label, fontsize=10, ha="left", va="bottom")

    ax.text(ROAD_END, 3.08, "107k", fontsize=10, ha="right", va="bottom")

    # IC 표시
    for name, x in IC_POINTS.items():
        ax.text(
            x,
            3.55,
            name,
            fontsize=10,
            fontweight="bold",
            color="blue",
            ha="center",
            va="bottom",
        )

    # 방향 라벨
    ax.text(-1.2, 1.5, "영암\n방향", fontsize=9, ha="right", va="center")
    ax.text(-1.2, -1.5, "순천\n방향", fontsize=9, ha="right", va="center")

    # 차로 라벨
    ax.text(108.0, 0.5, "1차로", fontsize=8, va="center")
    ax.text(108.0, 1.5, "2차로", fontsize=8, va="center")
    ax.text(108.0, 2.5, "갓길", fontsize=8, va="center")
    ax.text(108.0, -0.5, "1차로", fontsize=8, va="center")
    ax.text(108.0, -1.5, "2차로", fontsize=8, va="center")
    ax.text(108.0, -2.5, "갓길", fontsize=8, va="center")

    # 충돌 대상 ID
    conflict_unit_ids = set()

    if show_warnings and conflicts is not None and not conflicts.empty:
        for _, c in conflicts.iterrows():
            conflict_unit_ids.add(c["unit_id1"])
            conflict_unit_ids.add(c["unit_id2"])

    # 작업 박스
    for _, row in units.iterrows():
        lane_range = get_lane_y_range(row["방향"], row["차로"])

        if lane_range is None:
            continue

        y0, y1 = lane_range

        x0 = row["시점"]
        width = row["종점"] - row["시점"]
        height = y1 - y0

        is_warning = show_warnings and row["unit_id"] in conflict_unit_ids

        edge_color = "red" if is_warning else "black"
        line_width = 2.0 if is_warning else 1.0

        rect = Rectangle(
            (x0, y0),
            width,
            height,
            facecolor="#BFBFBF",
            edgecolor=edge_color,
            linewidth=line_width,
            zorder=3,
        )

        ax.add_patch(rect)

        # 박스 안 텍스트
        if row["다공종여부"]:
            if width >= 8:
                label = f"{row['그룹명']}\n{row['번호표시']}\n다공종"
                fontsize = 7
            else:
                label = f"{row['그룹명']}\n{row['번호표시']}"
                fontsize = 7
        else:
            if width >= 8 and not submit_mode:
                label = f"{row['번호표시']}\n{row['공사명']}"
                fontsize = 7
            else:
                label = f"{row['번호표시']}"
                fontsize = 9

        ax.text(
            x0 + width / 2,
            y0 + height / 2,
            label,
            fontsize=fontsize,
            ha="center",
            va="center",
            zorder=4,
        )

    if show_warnings:
        ax.text(
            0,
            -3.55,
            "빨간 음영: 구간 겹침 / 주황 음영: 기준 거리 이내 인접 / 빨간 테두리: 검토 대상 작업",
            fontsize=9,
            ha="left",
            va="center",
        )

    return fig


# =========================================================
# 12. Streamlit 화면
# =========================================================
st.set_page_config(
    page_title="보성지사 공사구간 도식 생성기 V2",
    layout="wide",
)

st.title("보성지사 공사구간 도식 생성기 V2")
st.caption("엑셀 업로드 기반 / 영암순천선 0k ~ 106.84k / 다공종 그룹 처리 / 검토용·제출용 출력")

with st.sidebar:
    st.header("설정")

    output_mode = st.radio(
        "출력 모드",
        ["검토용", "제출용"],
        index=0,
    )

    threshold = st.number_input(
        "인접 판단 거리(km)",
        min_value=0.0,
        max_value=20.0,
        value=5.0,
        step=0.5,
    )

    use_group = st.checkbox(
        "그룹명 기준으로 다공종 작업 묶기",
        value=True,
    )

    same_direction_only = st.checkbox(
        "같은 방향끼리만 검토",
        value=True,
    )

    consider_lane = st.checkbox(
        "차로까지 고려해서 검토",
        value=False,
    )

    st.markdown("---")
    st.subheader("자동 미표시 조건")

    hide_full_range = st.checkbox(
        "0~106.84 전체구간 기본 미표시",
        value=True,
    )

    hide_moving_closure = st.checkbox(
        "이동차단 작업 기본 미표시",
        value=True,
    )

    hide_no_lane = st.checkbox(
        "차로정보 없는 작업 기본 미표시",
        value=True,
    )

    st.markdown("---")
    st.markdown("""
    **방향 변환**
    - 순천종점 → 순천방향
    - 영암기점 → 영암방향
    - 양방향 → 순천 + 영암 둘 다 생성

    **이정 처리**
    - 공사구간에서 숫자 2개 추출
    - 방향과 관계없이 작은 km ~ 큰 km로 도식화

    **제출용**
    - 경고 음영과 빨간 테두리 숨김
    """)


st.subheader("1. 엑셀 업로드")

uploaded_file = st.file_uploader(
    "작업계획 엑셀 파일을 업로드하세요. (.xls / .xlsx)",
    type=["xls", "xlsx"],
)

parsed_df = pd.DataFrame()

if uploaded_file is not None:
    uploaded_file.seek(0)
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("시트 선택", excel_file.sheet_names)

    raw_df = read_excel_smart(uploaded_file, sheet_name)

    if raw_df.empty:
        st.error("엑셀에서 공사명/방향/공사구간/차단차로 헤더를 찾지 못했습니다.")
        st.stop()

    st.markdown("#### 원본 엑셀 미리보기")
    st.dataframe(raw_df.head(30), use_container_width=True)

    columns = list(raw_df.columns)

    default_name_col = guess_column(columns, ["공사명", "공사", "내용"])
    default_direction_col = guess_column(columns, ["방향"])
    default_section_col = guess_column(columns, ["공사구간", "구간", "이정"])
    default_lane_col = guess_column(columns, ["차단차로", "차로"])

    st.markdown("#### 컬럼 매칭")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        name_col = st.selectbox(
            "공사명 컬럼",
            columns,
            index=columns.index(default_name_col),
        )

    with col2:
        direction_col = st.selectbox(
            "방향 컬럼",
            columns,
            index=columns.index(default_direction_col),
        )

    with col3:
        section_col = st.selectbox(
            "공사구간 컬럼",
            columns,
            index=columns.index(default_section_col),
        )

    with col4:
        lane_col = st.selectbox(
            "차단차로 컬럼",
            columns,
            index=columns.index(default_lane_col),
        )

    parsed_df = parse_excel_to_work_table(
        raw_df,
        name_col=name_col,
        direction_col=direction_col,
        section_col=section_col,
        lane_col=lane_col,
        hide_full_range=hide_full_range,
        hide_moving_closure=hide_moving_closure,
        hide_no_lane=hide_no_lane,
    )

else:
    st.info("엑셀 파일을 업로드하면 자동으로 공사구간을 추출합니다.")

    parsed_df = pd.DataFrame([
        {
            "표시여부": True,
            "제외사유": "",
            "번호": 1,
            "공사명": "해룡교 교통통제 및 교량점검차 조사",
            "방향": "순천",
            "시점": 103.18,
            "종점": 104.87,
            "차로": "2차로",
            "그룹명": "",
            "원본행": 1,
            "원문방향": "순천종점",
            "원문공사구간": "103.18km ~ 104.87km",
            "원문차단차로": "2차로중 2차로",
        },
        {
            "표시여부": False,
            "제외사유": "이동차단",
            "번호": 2,
            "공사명": "단터널 소화기 이동식차단 점검",
            "방향": "순천",
            "시점": 2.3,
            "종점": 95.0,
            "차로": "2차로",
            "그룹명": "",
            "원본행": 2,
            "원문방향": "순천종점",
            "원문공사구간": "2.3km ~ 95km",
            "원문차단차로": "2차로중 2차로, 이동차단",
        },
        {
            "표시여부": False,
            "제외사유": "전체구간,이동차단",
            "번호": 3,
            "공사명": "일상유지보수 작업",
            "방향": "순천",
            "시점": 0.0,
            "종점": 106.84,
            "차로": "1차로,2차로,갓길",
            "그룹명": "",
            "원본행": 3,
            "원문방향": "양방향",
            "원문공사구간": "0km ~ 106.84km",
            "원문차단차로": "2차로중 1차로, 2차로, 이동차단, 갓길",
        },
        {
            "표시여부": False,
            "제외사유": "전체구간,이동차단",
            "번호": 4,
            "공사명": "일상유지보수 작업",
            "방향": "영암",
            "시점": 0.0,
            "종점": 106.84,
            "차로": "1차로,2차로,갓길",
            "그룹명": "",
            "원본행": 3,
            "원문방향": "양방향",
            "원문공사구간": "0km ~ 106.84km",
            "원문차단차로": "2차로중 1차로, 2차로, 이동차단, 갓길",
        },
        {
            "표시여부": True,
            "제외사유": "",
            "번호": 5,
            "공사명": "정밀안전점검 및 성능평가 용역",
            "방향": "영암",
            "시점": 49.3,
            "종점": 62.7,
            "차로": "2차로,갓길",
            "그룹명": "",
            "원본행": 5,
            "원문방향": "영암기점",
            "원문공사구간": "49.3km ~ 62.7km",
            "원문차단차로": "2차로중 2차로, 갓길",
        },
    ])


st.subheader("2. 자동 추출 결과 수정")

if parsed_df.empty:
    st.warning("추출된 공사구간이 없습니다. 컬럼 매칭 또는 원본 엑셀 내용을 확인하세요.")
else:
    total_count = len(parsed_df)
    display_count = int(parsed_df["표시여부"].sum()) if "표시여부" in parsed_df.columns else 0
    hidden_count = total_count - display_count

    m1, m2, m3 = st.columns(3)
    m1.metric("추출 작업 수", total_count)
    m2.metric("기본 표시", display_count)
    m3.metric("기본 미표시", hidden_count)

    search_text = st.text_input(
        "공사명 검색",
        value="",
        placeholder="공사가 많을 때 공사명을 입력해서 검색하세요.",
    )

    filter_option = st.radio(
        "표시 필터",
        ["전체", "표시 대상만", "미표시 대상만"],
        horizontal=True,
    )

    view_df = parsed_df.copy()

    if search_text.strip() != "":
        mask = view_df["공사명"].astype(str).str.contains(search_text.strip(), case=False, na=False)
        view_df = view_df[mask]

    if filter_option == "표시 대상만":
        view_df = view_df[view_df["표시여부"] == True]
    elif filter_option == "미표시 대상만":
        view_df = view_df[view_df["표시여부"] == False]

    st.caption("아래 표에서 표시여부를 끄면 도식에서 제외됩니다. 그룹명을 입력하면 다공종 작업으로 묶입니다.")

    edited_view_df = st.data_editor(
        view_df,
        num_rows="dynamic",
        use_container_width=True,
        height=420,
        column_config={
            "표시여부": st.column_config.CheckboxColumn(
                "표시",
                help="도식에 표시할 공사만 체크하세요.",
                default=True,
            ),
            "제외사유": st.column_config.TextColumn("기본 제외사유"),
            "번호": st.column_config.NumberColumn("번호", min_value=1, step=1),
            "공사명": st.column_config.TextColumn("공사명"),
            "방향": st.column_config.SelectboxColumn("방향", options=["순천", "영암"]),
            "시점": st.column_config.NumberColumn("시점(km)", min_value=0.0, max_value=ROAD_END, step=0.1),
            "종점": st.column_config.NumberColumn("종점(km)", min_value=0.0, max_value=ROAD_END, step=0.1),
            "차로": st.column_config.TextColumn("차로 예: 1차로 / 2차로 / 갓길 / 1차로,2차로"),
            "그룹명": st.column_config.TextColumn("다공종 그룹명 예: A, B"),
            "원본행": st.column_config.NumberColumn("원본행"),
            "원문방향": st.column_config.TextColumn("원문방향"),
            "원문공사구간": st.column_config.TextColumn("원문공사구간"),
            "원문차단차로": st.column_config.TextColumn("원문차단차로"),
        },
        key="edited_view_df",
    )

    # 필터링된 화면에서 수정한 내용을 원본 parsed_df에 반영
    parsed_df.update(edited_view_df)

    work_df = normalize_work_table(parsed_df)

    units_df = build_work_units(
        work_df,
        use_group=use_group,
    )

    show_warnings = output_mode == "검토용"

    st.subheader("3. 도식 반영 대상")

    if work_df.empty:
        st.warning("표시 대상으로 선택된 공사가 없습니다.")
    else:
        st.dataframe(
            work_df[["번호", "공사명", "방향", "시점", "종점", "차로표시", "그룹명"]],
            use_container_width=True,
            height=260,
        )

        st.subheader("4. 다공종 묶음 결과")

        if units_df.empty:
            st.warning("도식화할 작업 단위가 없습니다.")
        else:
            st.dataframe(
                units_df[[
                    "번호표시",
                    "공사명",
                    "상세공사명",
                    "방향",
                    "시점",
                    "종점",
                    "차로표시",
                    "그룹명",
                    "다공종여부",
                ]],
                use_container_width=True,
                height=260,
            )

            conflicts = find_conflicts(
                units_df,
                threshold_km=threshold,
                same_direction_only=same_direction_only,
                consider_lane=consider_lane,
            )

            st.subheader("5. 겹침 / 인접 판정")

            if output_mode == "검토용":
                if conflicts.empty:
                    st.success("겹치는 구간 또는 기준 거리 이내 인접 구간이 없습니다.")
                else:
                    st.error(f"주의가 필요한 구간이 {len(conflicts)}건 확인되었습니다.")

                    st.dataframe(
                        conflicts[[
                            "작업1",
                            "작업2",
                            "방향",
                            "구분",
                            "문제구간",
                            "이격거리(km)",
                        ]],
                        use_container_width=True,
                        height=260,
                    )
            else:
                st.info("제출용 모드입니다. 경고 음영과 빨간 테두리는 표시하지 않습니다.")

            st.subheader("6. 공사구간 도식")

            fig = draw_diagram(
                units_df,
                conflicts,
                show_warnings=show_warnings,
                submit_mode=(output_mode == "제출용"),
            )

            st.pyplot(fig, use_container_width=True)

            buffer = io.BytesIO()
            fig.savefig(buffer, format="png", bbox_inches="tight")
            buffer.seek(0)

            file_name = (
                "bosung_work_diagram_v2_review.png"
                if output_mode == "검토용"
                else "bosung_work_diagram_v2_submit.png"
            )

            st.download_button(
                label="PNG 이미지 다운로드",
                data=buffer,
                file_name=file_name,
                mime="image/png",
            )
