from __future__ import annotations

import io
import math
import re
import datetime as dt
from copy import copy
from dataclasses import dataclass
from typing import List, Tuple, Dict, Optional

import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side

# =========================
# 상수/스타일
# =========================
PRODUCT_NAME_CANDIDATES = ["등록상품명", "상품명", "노출상품명", "쿠팡노출상품명"]
OPTION_NAME_COL = "등록 옵션명"

RE_SIZE = re.compile(r"(?<!\d)(\d{3})(?!\d)")
WS = re.compile(r"\s+")
COLOR_LONG2SHORT = {
    "블랙": "흑", "화이트": "백", "아이보리": "아이보리",
    "그레이": "회", "베이지": "베", "네이비": "곤",
    "핑크": "핑", "레드": "적", "브라운": "밤",
    "실버": "실", "골드": "금",
}
SHORT_COLORS = {"흑","백","아이보리","회","베","곤","핑","적","밤","실","금"}
COLOR_WORDS = list(COLOR_LONG2SHORT.keys()) + list(SHORT_COLORS)
BRAND_TITLE = {"준디": "준디자인", "자유": "자유디자인"}

BRIGHT_YELLOW = PatternFill("solid", fgColor="FFFF00")
THIN   = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# =========================
# 공통 유틸
# =========================
def normalize_text(s: str) -> str:
    if pd.isna(s):
        return ""
    return WS.sub(" ", str(s)).strip()

def find_product_name_col(df: pd.DataFrame) -> Optional[str]:
    for c in PRODUCT_NAME_CANDIDATES:
        if c in df.columns:
            return c
    stripped = {col.replace(" ", ""): col for col in df.columns}
    for c in PRODUCT_NAME_CANDIDATES:
        cc = c.replace(" ", "")
        if cc in stripped:
            return stripped[cc]
    return None

def shorten_color(token: str) -> str:
    if not token:
        return token
    t = str(token).strip()
    return COLOR_LONG2SHORT.get(t, t)

def _today_md() -> str:
    now = dt.datetime.now()
    return f"{now.month}/{now.day}"

def _clean_value(v):
    if v is None:
        return None
    s = str(v).strip()
    if s == "" or s.lower() in {"none", "nan"}:
        return None
    return v

def _write(ws, row, col, v, as_int=False):
    v = _clean_value(v)
    if v is None:
        ws.cell(row=row, column=col, value=None)
        return
    if as_int:
        try:
            ws.cell(row=row, column=col, value=int(v)); return
        except Exception:
            pass
    ws.cell(row=row, column=col, value=v)

# =========================
# 옵션 파싱
# =========================
def _preprocess_option_text(text: str) -> str:
    if not text:
        return ""
    s = normalize_text(text)
    s = re.sub(r'\b(색상|사이즈)\b\s*[:：]?', ' ', s)
    color_alt = "|".join(map(re.escape, COLOR_WORDS))
    s = re.sub(rf'(여성|남성)\s*({color_alt})', r'\1 \2', s)
    s = re.sub(r'(여성|남성)(\d{3})', r'\1 \2', s)
    return WS.sub(" ", s).strip()

def _find_color_near_size(s: str, size_span: Tuple[int,int]) -> Optional[str]:
    color_regex = re.compile("|".join(sorted(map(re.escape, COLOR_WORDS), key=len, reverse=True)))
    candidates = [(m.group(), m.start(), m.end()) for m in color_regex.finditer(s)]
    if not candidates:
        return None
    size_center = (size_span[0] + size_span[1]) / 2
    best = None; best_dist = 10**9
    for word, a, b in candidates:
        center = (a + b) / 2
        dist = abs(center - size_center)
        if dist < best_dist:
            best_dist = dist; best = word
    return best

def parse_color_size(option_name: str) -> Tuple[Optional[str], Optional[str]]:
    if not option_name:
        return None, None
    text = _preprocess_option_text(option_name)
    m = RE_SIZE.search(text)
    if not m:
        return None, None
    size = m.group(1)
    color_word = _find_color_near_size(text, m.span())
    if not color_word:
        toks = text.split()
        for tok in reversed(toks):
            if tok in COLOR_WORDS:
                color_word = tok; break
    if not color_word:
        return None, None
    base_short = shorten_color(color_word)
    is_women = '여성' in text; is_men = '남성' in text
    if is_women:   color_label = f"{base_short}中"
    elif is_men:   color_label = f"{base_short}大"
    else:          color_label = base_short
    return color_label, size

def extract_colors_sizes(option_series: pd.Series) -> Tuple[List[str], List[str]]:
    colors, sizes = [], []; seen_c, seen_s = set(), set()
    for v in option_series.dropna().astype(str):
        c, s = parse_color_size(v)
        if c and c not in seen_c: colors.append(c); seen_c.add(c)
        if s and s not in seen_s: sizes.append(s); seen_s.add(s)
    return colors, sizes

# =========================
# 세션/그리드
# =========================
def init_session():
    st.session_state.setdefault("basket", [])
    st.session_state.setdefault("current_grid", pd.DataFrame())
    st.session_state.setdefault("current_context", {})
    st.session_state.setdefault("manual_product", "")
    st.session_state.setdefault("manual_colors_text", "")
    st.session_state.setdefault("manual_size_start", 230)
    st.session_state.setdefault("manual_size_end", 280)
    st.session_state.setdefault("manual_size_step", 5)

def add_grid_to_basket(grid_df: pd.DataFrame, brand: str, product: str):
    for color in grid_df.index:
        for size in grid_df.columns:
            val = grid_df.loc[color, size]
            if pd.isna(val): continue
            try:
                q = int(val)
            except:
                continue
            if q <= 0: continue
            st.session_state["basket"].append({
                "brand": brand, "product": product,
                "color": str(color), "size": str(size), "qty": q
            })

def clear_basket():
    st.session_state["basket"] = []

def build_grid(colors: List[str], sizes: List[str], selectable_colors: List[str]) -> pd.DataFrame:
    if not sizes: return pd.DataFrame()
    if not selectable_colors: selectable_colors = []
    df = pd.DataFrame(index=selectable_colors, columns=sizes, data=math.nan)
    df.index.name = "색상"
    return df

# =========================
# 템플릿 조작
# =========================
def _set_header(ws, brand: str):
    header_txt = f"{_today_md()} {BRAND_TITLE.get(brand, brand)}(&P번)"
    try:
        ws.header_footer.center_header = header_txt
    except Exception:
        try:
            ws.oddHeader.center.text = header_txt
            ws.evenHeader.center.text = header_txt
        except Exception:
            pass
    try:
        ws.page_setup.firstPageNumber = 51
        ws.page_setup.use_firstPageNumber = True
    except Exception:
        pass

def _clear_data_area_keep_row1(ws, start_row=2, start_col=2):
    maxr = ws.max_row or 2000
    maxc = ws.max_column or 200
    for r in range(start_row, maxr + 1):
        for c in range(start_col, maxc + 1):
            ws.cell(row=r, column=c).value = None

def export_brand_template_xlsx(brand: str, items: List[dict], template_bytes: bytes) -> bytes:
    """
    - 업로드한 template.xlsx의 1행/서식 보존
    - B2부터 기록, 제품명은 FFFF00 하이라이트
    - 제품별 '실사용 사이즈'만 가로 헤더(C~) 생성, C열 이후 너비=4
    """
    try:
        wb = load_workbook(io.BytesIO(template_bytes))
    except Exception:
        wb = Workbook()
    ws = wb.active
    _set_header(ws, brand)
    _clear_data_area_keep_row1(ws, start_row=2, start_col=2)

    by_product: Dict[str, List[dict]] = {}
    for it in items:
        by_product.setdefault(it["product"], []).append(it)

    row = 2
    for product_raw, p_items in by_product.items():
        # 실제 사용 사이즈 수집
        sizes_used, seen = [], set()
        for it in p_items:
            try:
                q = int(it["qty"])
            except:
                continue
            if q > 0:
                sz = str(it["size"])
                if sz not in seen:
                    seen.add(sz); sizes_used.append(sz)

        def _size_key(x):
            m = re.search(r"\d+", str(x))
            return (0, int(m.group())) if m else (1, str(x))
        sizes_used.sort(key=_size_key)

        # 색상별 합계
        by_color: Dict[str, Dict[str, int]] = {}
        for it in p_items:
            try:
                q = int(it["qty"])
            except:
                continue
            if q <= 0: continue
            color = str(it["color"]); sz = str(it["size"])
            by_color.setdefault(color, {})
            by_color[color][sz] = by_color[color].get(sz, 0) + q

        # 제품명
        _write(ws, row, 2, product_raw)
        ws.cell(row=row, column=2).fill = BRIGHT_YELLOW

        # 사이즈 헤더
        start_col = 3
        for i, sz in enumerate(sizes_used):
            col = start_col + i
            _write(ws, row, col, sz)
            ws.column_dimensions[get_column_letter(col)].width = 4

        # 색상 행
        for color, size_map in by_color.items():
            row += 1
            _write(ws, row, 2, color)
            for i, sz in enumerate(sizes_used):
                col = start_col + i
                val = size_map.get(sz)
                _write(ws, row, col, val, as_int=True)

        row += 1

    # 마무리: C~ 마지막까지 너비=4
    maxc = ws.max_column or 3
    for c in range(3, maxc + 1):
        ws.column_dimensions[get_column_letter(c)].width = 4

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================
# 데이터 로드 & 검색
# =========================
def load_uploaded_df(file) -> Optional[pd.DataFrame]:
    try:
        df = pd.read_excel(file)
        if OPTION_NAME_COL not in df.columns:
            st.error(f"엑셀에 '{OPTION_NAME_COL}' 컬럼이 필요합니다.")
            return None
        name_col = find_product_name_col(df)
        if not name_col:
            st.error(f"제품명 컬럼이 필요합니다. 후보: {', '.join(PRODUCT_NAME_CANDIDATES)}")
            return None
        df[name_col] = df[name_col].map(normalize_text)
        df[OPTION_NAME_COL] = df[OPTION_NAME_COL].map(normalize_text)
        return df
    except Exception as e:
        st.exception(e); return None

def search_products(df: pd.DataFrame, keyword: str, name_col: str) -> pd.DataFrame:
    if not keyword:
        return df[[name_col]].drop_duplicates().head(200)
    kw = keyword.strip()
    mask = df[name_col].astype(str).str.contains(re.escape(kw), case=False, na=False)
    return df.loc[mask, [name_col]].drop_duplicates().head(200)

def _sizes_range_step(start: int, end: int, step: int) -> List[str]:
    if step <= 0 or end < start: return []
    return [str(s) for s in range(start, end + 1, step)]

# =========================
# 탭 본문
# =========================
def tab_body(brand: str, df: pd.DataFrame, template_bytes: bytes):
    name_col = find_product_name_col(df)

    st.subheader(f"{brand} — 제품명 검색")
    q = st.text_input("제품명 검색 (부분 검색)", key=f"q_{brand}")
    result = search_products(df, q, name_col)
    st.caption(f"검색 결과 {len(result)}개 (최대 200개 표시)")

    product = st.selectbox(
        "제품 선택",
        ["(선택)"] + result[name_col].tolist(),
        key=f"product_{brand}"
    )
    manual_toggle = st.checkbox("직접입력 모드 (파싱 무시하고 수동 작성)", key=f"manual_{brand}")

    # 자동 파싱
    if product != "(선택)" and not manual_toggle:
        sub = df[df[name_col] == product]
        colors, sizes = extract_colors_sizes(sub[OPTION_NAME_COL])

        if not sizes:
            st.warning("이 제품은 '등록 옵션명'에서 사이즈를 찾지 못했습니다. 아래 '직접입력 모드'를 사용하세요.")
            manual_toggle = True

        if not manual_toggle:
            st.write("---")
            st.subheader("2) 표 입력 (세로=색상 / 가로=사이즈)")
            color_choices = st.multiselect(
                "세로(행)에 표시할 색상 선택",
                options=colors,
                default=colors[: min(5, len(colors))],
                key=f"colors_{brand}"
            )

            grid = build_grid(colors, sizes, color_choices)
            st.caption("각 칸에 수량을 입력하세요. 비우거나 0은 미반영됩니다.")
            edited = st.data_editor(
                grid, key=f"editor_{brand}",
                use_container_width=True, num_rows="fixed",
            )

            c1, c2 = st.columns(2)
            with c1:
                if st.button("현재 표를 리스트업에 추가", key=f"add_{brand}"):
                    add_grid_to_basket(edited, brand=brand, product=product)
                    st.success("리스트업에 추가했습니다.")
            with c2:
                if st.button("리스트업 비우기", key=f"clear_{brand}"):
                    clear_basket(); st.info("리스트업을 비웠습니다.")

    # 직접 입력
    if manual_toggle:
        st.write("---"); st.subheader("직접입력 모드")
        mp = st.text_input("제품명(직접 입력)", value=(product if product != "(선택)" else st.session_state.get("manual_product", "")), key=f"mp_{brand}")
        st.session_state["manual_product"] = mp

        st.caption("자주 쓰는 색상 프리셋")
        preset_colors = ["흑","백","곤","핑","회","베","적","아이보리"]
        cols = st.columns(len(preset_colors))
        for i, col in enumerate(cols):
            if col.button(preset_colors[i], key=f"preset_{brand}_{i}"):
                mc_text = st.session_state.get("manual_colors_text", "")
                parts = [p.strip() for p in mc_text.split(",") if p.strip()]
                if preset_colors[i] not in parts: parts.append(preset_colors[i])
                st.session_state["manual_colors_text"] = ",".join(parts)

        mc_text = st.text_input(
            "색상 목록(쉼표로 구분, 예: 흑, 백, 아이보리)",
            value=st.session_state.get("manual_colors_text", ""),
            key=f"mc_{brand}"
        )
        st.session_state["manual_colors_text"] = mc_text
        colors = [c.strip() for c in mc_text.split(",") if c.strip()]

        c1, c2, c3 = st.columns(3)
        with c1:
            start = st.number_input("시작 사이즈", min_value=200, max_value=400, value=st.session_state.get("manual_size_start", 230), step=5, key=f"start_{brand}")
        with c2:
            end = st.number_input("끝 사이즈", min_value=200, max_value=420, value=st.session_state.get("manual_size_end", 280), step=5, key=f"end_{brand}")
        with c3:
            step = st.number_input("간격", min_value=1, max_value=50, value=st.session_state.get("manual_size_step", 5), step=1, key=f"step_{brand}")

        st.session_state["manual_size_start"] = int(start)
        st.session_state["manual_size_end"] = int(end)
        st.session_state["manual_size_step"] = int(step)

        sizes = _sizes_range_step(int(start), int(end), int(step))

        if not mp:
            st.info("제품명을 입력하면 표가 생성됩니다."); return
        if not colors:
            st.info("색상을 한 개 이상 입력하세요. (쉼표로 구분)"); return
        if not sizes:
            st.warning("사이즈 범위/간격을 다시 확인하세요."); return

        st.subheader(f"2) 표 입력 (세로=색상 / 가로=사이즈 {start}~{end}, {step} 단위)")
        grid = build_grid(colors, sizes, colors)
        st.caption("각 칸에 수량을 입력하세요. 비우거나 0은 미반영됩니다.")
        edited = st.data_editor(
            grid, key=f"editor_manual_{brand}",
            use_container_width=True, num_rows="fixed",
        )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("현재 표를 리스트업에 추가(직접입력)", key=f"add_manual_{brand}"):
                add_grid_to_basket(edited, brand=brand, product=mp); st.success("리스트업에 추가했습니다.")
        with c2:
            if st.button("리스트업 비우기", key=f"clear_manual_{brand}"):
                clear_basket(); st.info("리스트업을 비웠습니다.")

    # 현재 브랜드만 보기/다운로드
    basket = [e for e in st.session_state["basket"] if e["brand"] == brand]
    st.write("---"); st.subheader(f"{brand} 누적 리스트업 미리보기")
    if not basket:
        st.caption("아직 추가된 항목이 없습니다.")
    else:
        st.dataframe(pd.DataFrame(basket), use_container_width=True, height=280)
        xlsx_bytes = export_brand_template_xlsx(brand, basket, st.session_state["template_bytes"])
        fname = f"{brand}_{dt.datetime.now():%Y%m%d}_template.xlsx"
        st.download_button(
            f"{brand} 템플릿 다운로드 (xlsx)",
            data=xlsx_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# =========================
# 메인
# =========================
def main():
    st.set_page_config(page_title="브랜드별 템플릿 빌더(클라우드)", layout="wide")
    init_session()

    st.title("TXT 없이 바로 템플릿 만들기 (브랜드별 파일, 클라우드 업로드 버전)")
    st.caption("템플릿 보존 · B2부터 기록 · 제품명 하이라이트=FFFF00 · 제품별 가변 사이즈 헤더 · C열부터 너비=4 고정")

    with st.expander("업로드 안내", expanded=True):
        st.markdown(
            """
            1) **template.xlsx**: 출력 서식/머리글/인쇄설정으로 사용할 템플릿 파일  
            2) **준디 상품리스트.xlsx**, **자유 상품리스트.xlsx**: 각 파일에 **'등록 옵션명'** 컬럼이 있어야 합니다.  
            - 제품명 컬럼은 자동 인식(후보: 등록상품명/상품명/노출상품명/쿠팡노출상품명)
            - 색상 표기는 자동 치환: 블랙→흑, 화이트→백, 여성→中, 남성→大
            """
        )

    c1, c2, c3 = st.columns(3)
    with c1:
        template_file = st.file_uploader("template.xlsx 업로드", type=["xlsx"], key="tpl")
    with c2:
        jundi_file = st.file_uploader("준디 상품리스트 업로드", type=["xlsx"], key="jundi")
    with c3:
        jayu_file  = st.file_uploader("자유 상품리스트 업로드", type=["xlsx"], key="jayu")

    if not template_file:
        st.info("template.xlsx를 업로드해 주세요."); return
    st.session_state["template_bytes"] = template_file.read()

    tabs = []
    if jundi_file:
        df_jundi = load_uploaded_df(jundi_file)
        if df_jundi is not None:
            tabs.append(("준디", df_jundi))
    if jayu_file:
        df_jayu = load_uploaded_df(jayu_file)
        if df_jayu is not None:
            tabs.append(("자유", df_jayu))

    if not tabs:
        st.warning("좌측 파일 업로드 후 탭이 나타납니다."); return

    st.write("---")
    t_objs = st.tabs([t[0] for t in tabs])
    for tab, (brand, df) in zip(t_objs, tabs):
        with tab:
            tab_body(brand, df, st.session_state["template_bytes"])

if __name__ == "__main__":
    main()
