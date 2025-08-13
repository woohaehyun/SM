
import streamlit as st
import pandas as pd
import io
import zipfile
import os
import re
from datetime import timedelta

st.set_page_config(page_title="신명약품 자동발주", layout="wide")

# ============= 사이드바 / 헤더 =============
st.sidebar.header("📂 파일 업로드")
sales_file = st.sidebar.file_uploader("매출자료 업로드", type=["xlsx"])
purchase_file = st.sidebar.file_uploader("매입자료 업로드", type=["xlsx"])
stock_file = st.sidebar.file_uploader("현재고 업로드", type=["xlsx"])
logo_upload = st.sidebar.file_uploader("로고 이미지(선택)", type=["png","jpg","jpeg","webp"])
name_map_file = st.sidebar.file_uploader("명칭 매핑표(선택: from,to)", type=["xlsx","csv"])

st.sidebar.divider()
# 발주서 양식 옵션: 매입처 제외(제조사 기준)
manufacturer_only = st.sidebar.checkbox("발주서 양식: 매입처 제외(제조사 기준)", value=True)

# (고급) 내부 표시는 자유롭게 보되, 실제 내보내기는 제조사 기준으로 강제
group_mode = st.sidebar.radio("📁 화면 그룹 기준(미리보기용)", ["제조사", "매입처", "제조사+매입처"], index=0, help="실제 발주서 파일은 위 체크가 켜져 있으면 제조사 기준으로만 만들어집니다.")

st.sidebar.divider()
mode = st.sidebar.radio("📅 분석 기간", ["자동 (최근 3개월)", "수동 지정"])

st.sidebar.divider()
use_recent_purchase = st.sidebar.checkbox("최근 입고수량 반영하여 과발주 방지", value=True)
recent_days = st.sidebar.number_input("최근 입고 반영 일수", min_value=1, max_value=90, value=14, step=1)

st.sidebar.divider()
days_options = list(range(1, 366))
days_label_map = {d: f"{d}일" for d in days_options}
selected_days = st.sidebar.selectbox("발주 기준(최근 N일 판매량)", options=days_options, format_func=lambda x: days_label_map[x], index=29)  # 기본 30일

min_shortage = st.sidebar.number_input("부족수량 하한(이상만 표시)", min_value=0, value=0, step=1)
show_only_to_order = st.sidebar.checkbox("발주 필요 항목만 보기(부족수량>0)", value=True)

st.sidebar.divider()
export_mode = st.sidebar.radio("엑셀 내보내기 방식", ["그룹별 개별 파일 (ZIP)", "한 파일(탭 구분)"], index=1)

# ===== 헤더 영역 =====
c1, c2 = st.columns([1, 5])
with c1:
    try:
        if logo_upload is not None:
            st.image(logo_upload, width=100)
        elif os.path.exists("로고리뉴얼.png"):
            st.image("로고리뉴얼.png", width=230)
        else:
            st.empty()
    except Exception:
        st.empty()
with c2:
    st.title("💊 신명약품 자동발주앱")
   
# ======== 유틸 ========
def normalize_columns(df, mapping):
    df = df.copy()
    df.rename(columns={k: v for k, v in mapping.items() if k in df.columns}, inplace=True)
    return df

def require_columns(df, required, name):
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"{name}에 필요한 컬럼이 없습니다: {', '.join(missing)}")
        st.stop()

def to_upper_strip(series):
    return series.astype(str).str.strip().str.upper()

def clean_party_name(series):
    rep = [("㈜",""),("(주)",""),("주식회사",""),("(유)",""),("유한회사",""),("(재)",""),("재단법인",""),("(사)",""),("사단법인","")]
    s = series.astype(str).str.upper()
    for a,b in rep:
        s = s.str.replace(a, b, regex=False)
    s = s.str.replace(r"\s+", " ", regex=True).str.strip()
    return s

def apply_name_mapping(series, mapping_df):
    if mapping_df is None:
        return series
    m = mapping_df.copy()
    m.columns = [c.strip().lower() for c in m.columns]
    if not set(["from","to"]).issubset(set(m.columns)):
        return series
    s_clean = series.astype(str).str.upper().str.strip()
    m["from"] = m["from"].astype(str).str.upper().str.strip()
    m["to"] = m["to"].astype(str).str.strip()
    return s_clean.map(dict(zip(m["from"], m["to"]))).fillna(series)

def sanitize_sheet_name(name: str) -> str:
    if name is None or str(name).strip() == "":
        return "미지정"
    s = str(name)
    s = re.sub(r"[\[\]\*\?/\\:]", "-", s)
    return s[:31]

def write_formatted_sheet(writer, sheet_name, df):
    df = df.copy()
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    workbook  = writer.book
    ws = writer.sheets[sheet_name]

    header_fmt = workbook.add_format({"bold": True, "bg_color": "#F2F4F7", "border": 1, "align": "center", "valign": "vcenter"})
    num_fmt    = workbook.add_format({"num_format": "#,##0"})
    strong_fmt = workbook.add_format({"bg_color": "#FFE5E5", "bold": True})
    over_fmt   = workbook.add_format({"bg_color": "#EAF4FF"})
    base_fmt   = workbook.add_format({"text_wrap": False})

    # 헤더 서식
    for col_idx, col in enumerate(df.columns):
        ws.write(0, col_idx, col, header_fmt)

    # 열 너비 자동
    for i, col in enumerate(df.columns):
        try:
            maxlen = max(df[col].astype(str).map(len).max(), len(col))
        except Exception:
            maxlen = len(col)
        ws.set_column(i, i, min(maxlen + 2, 40), num_fmt if pd.api.types.is_numeric_dtype(df[col]) else base_fmt)

    ws.set_default_row(20)
    ws.set_row(0, 24)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(df), len(df.columns)-1)

    # 조건부서식
    col_map = {c:i for i,c in enumerate(df.columns)}
    last_row = len(df)
    if "부족수량" in col_map:
        i = col_map["부족수량"]
        ws.conditional_format(1, i, last_row, i, {"type":"cell","criteria":">","value":0,"format":strong_fmt})
    if "발주수량" in col_map:
        i = col_map["발주수량"]
        ws.conditional_format(1, i, last_row, i, {"type":"cell","criteria":">","value":0,"format":strong_fmt})
    if "과재고" in col_map:
        i = col_map["과재고"]
        ws.conditional_format(1, i, last_row, i, {"type":"cell","criteria":">","value":0,"format":over_fmt})

# ==== 데이터 검증 리포트(간단 요약) ====
def validate_and_report(sales_raw, purchase_raw, stock_raw):
    req_sales = ["명세일자", "상 품 명", "포장단위", "수량"]
    req_purch = ["입고일자", "상 품 명", "포장단위", "수량"]
    req_stock = ["상 품 명", "포장단위", "재고수량"]

    missing_sales = [c for c in req_sales if c not in sales_raw.columns]
    missing_purch = [c for c in req_purch if c not in purchase_raw.columns]
    missing_stock = [c for c in req_stock if c not in stock_raw.columns]

    sales_date = pd.to_datetime(sales_raw.get("명세일자"), errors="coerce")
    purch_date = pd.to_datetime(purchase_raw.get("입고일자"), errors="coerce")
    summary = pd.DataFrame({
        "항목": ["필수컬럼누락(매출)","필수컬럼누락(매입)","필수컬럼누락(현재고)","날짜오류(매출)","날짜오류(매입)"],
        "건수": [len(missing_sales), len(missing_purch), len(missing_stock),
                int(sales_date.isna().sum()), int(purch_date.isna().sum())]
    })
    return summary

# ======== 메인 로직 ========
if sales_file and purchase_file and stock_file:
    sales_df = pd.read_excel(sales_file)  # 매출
    purchase_df = pd.read_excel(purchase_file)  # 매입(입고)
    stock_df = pd.read_excel(stock_file)  # 현재고

    # 컬럼 정규화(별칭 대응)
    sales_df = normalize_columns(sales_df, {
        "거래일자": "명세일자", "일자": "명세일자", "매출처": "매 출 처",
        "상품명": "상 품 명", "포장 단위": "포장단위", "제조사": "제 조 사", "제약사": "제 조 사",
        "공급처": "매 입 처", "거래처": "매 입 처", "매입처": "매 입 처"
    })
    purchase_df = normalize_columns(purchase_df, {
        "입고일": "입고일자", "거래처": "매 입 처", "매입처": "매 입 처", "공급처": "매 입 처",
        "상품명": "상 품 명", "포장 단위": "포장단위", "제조사": "제 조 사", "제약사": "제 조 사"
    })
    stock_df = normalize_columns(stock_df, {
        "거래처": "매 입 처", "매입처": "매 입 처", "공급처": "매 입 처",
        "상품명": "상 품 명", "포장 단위": "포장단위", "제조사": "제 조 사", "제약사": "제 조 사",
        "재고": "재고수량"
    })

    # 필수 컬럼 체크 (매입처는 필수가 아님)
    require_columns(sales_df, ["명세일자", "상 품 명", "포장단위", "수량"], "매출자료")
    require_columns(purchase_df, ["입고일자", "상 품 명", "포장단위", "수량"], "매입자료")
    require_columns(stock_df, ["상 품 명", "포장단위", "재고수량"], "현재고")

    # 문자열 정리
    for df in [sales_df, purchase_df, stock_df]:
        df["상 품 명"] = to_upper_strip(df["상 품 명"])
        df["포장단위"] = to_upper_strip(df["포장단위"])

    # 날짜형
    sales_df["명세일자"] = pd.to_datetime(sales_df["명세일자"], errors="coerce")
    purchase_df["입고일자"] = pd.to_datetime(purchase_df["입고일자"], errors="coerce")

    # 제조사/매입처 표준화 및 매핑 적용(매입처 없어도 무시)
    for df in [sales_df, purchase_df, stock_df]:
        if "제 조 사" in df.columns:
            df["제 조 사"] = clean_party_name(df["제 조 사"])
        if "매 입 처" in df.columns:
            df["매 입 처"] = clean_party_name(df["매 입 처"])
    # 매핑표 적용
    map_df = None
    if name_map_file is not None:
        try:
            map_df = pd.read_excel(name_map_file) if name_map_file.name.lower().endswith(".xlsx") else pd.read_csv(name_map_file)
        except Exception:
            map_df = None
    if map_df is not None:
        if "제 조 사" in stock_df.columns:
            stock_df["제 조 사"] = apply_name_mapping(stock_df["제 조 사"], map_df)
        if "매 입 처" in stock_df.columns:
            stock_df["매 입 처"] = apply_name_mapping(stock_df["매 입 처"], map_df)

    # 분석 기간
    if mode == "자동 (최근 3개월)":
        end_date = sales_df["명세일자"].max()
        start_date = end_date - pd.DateOffset(months=3)
    else:
        c3, c4 = st.columns(2)
        with c3:
            start_date = st.date_input("시작일", value=sales_df["명세일자"].min().date())
        with c4:
            end_date = st.date_input("종료일", value=sales_df["명세일자"].max().date())
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)

    # ===== 발주 기준: 최근 N일 판매량(총합) =====
    max_sale_date = sales_df["명세일자"].max()
    nday_start = max_sale_date - pd.Timedelta(days=int(selected_days))
    nday_sales = sales_df[(sales_df["명세일자"] > nday_start) & (sales_df["명세일자"] <= max_sale_date)]
    nday_qty = nday_sales.groupby(["상 품 명", "포장단위"], as_index=False)["수량"].sum()
    nday_qty.rename(columns={"수량": f"최근{selected_days}일_판매량"}, inplace=True)
    nday_qty["기준판매량"] = nday_qty[f"최근{selected_days}일_판매량"].astype(int)

    # 현재고 병합(제조사/매입처는 없을 수 있음)
    cols_to_pull = ["상 품 명", "포장단위", "재고수량"]
    if "제 조 사" in stock_df.columns: cols_to_pull.append("제 조 사")
    if "매 입 처" in stock_df.columns: cols_to_pull.append("매 입 처")
    merged = nday_qty.merge(stock_df[cols_to_pull].drop_duplicates(), on=["상 품 명", "포장단위"], how="left")

    # 제조사/매입처 보강: 최근 매입 이력에서 가져오기 (없으면 통과)
    purchase_sorted = purchase_df.sort_values("입고일자")
    agg_dict = {}
    if "제 조 사" in purchase_df.columns: agg_dict["제 조 사"] = "last"
    if "매 입 처" in purchase_df.columns: agg_dict["매 입 처"] = "last"
    if agg_dict:
        last_info = purchase_sorted.groupby(["상 품 명", "포장단위"]).agg(agg_dict).reset_index()
        merged = merged.merge(last_info, on=["상 품 명", "포장단위"], how="left", suffixes=("", "_최근입고"))
        for col in agg_dict.keys():
            if col in merged.columns and f"{col}_최근입고" in merged.columns:
                merged[col] = merged[col].fillna(merged[f"{col}_최근입고"])
                merged.drop(columns=[f"{col}_최근입고"], inplace=True)

    # 최근 입고 반영(옵션)
    if use_recent_purchase:
        cutoff = purchase_df["입고일자"].max() - pd.Timedelta(days=int(recent_days))
        recent_purchase = purchase_df[purchase_df["입고일자"] >= cutoff]
        recent_in_qty = recent_purchase.groupby(["상 품 명", "포장단위"], as_index=False)["수량"].sum()
        recent_in_qty.rename(columns={"수량": "최근입고수량"}, inplace=True)
        merged = merged.merge(recent_in_qty, on=["상 품 명", "포장단위"], how="left")
        merged["최근입고수량"] = merged["최근입고수량"].fillna(0).astype(int)
    else:
        merged["최근입고수량"] = 0

    # 부족/과재고/발주수량 계산
    merged["재고수량"] = merged["재고수량"].fillna(0).astype(int)
    merged["기준판매량"] = merged["기준판매량"].fillna(0).astype(int)

    merged["부족수량"] = (merged["기준판매량"] - merged["재고수량"] - merged["최근입고수량"]).astype(int)
    merged["부족수량"] = merged["부족수량"].apply(lambda x: x if x > 0 else 0)

    merged["과재고"] = (merged["재고수량"] - merged["기준판매량"]).astype(int)
    merged["과재고"] = merged["과재고"].apply(lambda x: x if x > 0 else 0)

    merged["발주수량"] = merged["부족수량"]

    # 보기 옵션 필터
    if min_shortage > 0:
        merged = merged[merged["부족수량"] >= int(min_shortage)]
    if show_only_to_order:
        merged = merged[merged["발주수량"] > 0]

    # ===== 상단 KPI =====
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("총 품목수", f"{len(merged):,}")
    k2.metric("발주 필요 품목수", f"{(merged['발주수량'] > 0).sum():,}")
    k3.metric("부족수량 합계", f"{int(merged['부족수량'].sum()):,}")
    k4.metric("과재고 합계", f"{int(merged['과재고'].sum()):,}")

    # ===== 미리보기(화면 전용 그룹 기준) =====
    f1, f2, f3 = st.columns([2, 1, 1])
    with f1:
        keyword = st.text_input("🔎 상품명 검색(대소문자 무시)", value="").strip().upper()
    with f2:
        manu_sel = st.multiselect("제조사 필터", sorted(pd.Series(merged.get("제 조 사", pd.Series())).dropna().unique().tolist())) if "제 조 사" in merged.columns else st.multiselect("제조사 필터", [])
    with f3:
        if "매 입 처" in merged.columns and not manufacturer_only:
            supplier_sel = st.multiselect("매입처 필터(미리보기)", sorted(pd.Series(merged["매 입 처"]).dropna().unique().tolist()))
        else:
            supplier_sel = []
            if "매 입 처" in merged.columns:
                st.caption("ℹ️ 발주서 양식에서 매입처는 제외됩니다.")

    view_df = merged.copy()
    if keyword:
        view_df = view_df[view_df["상 품 명"].str.contains(keyword, na=False)]
    if manu_sel and "제 조 사" in view_df.columns:
        view_df = view_df[view_df["제 조 사"].isin(manu_sel)]
    if supplier_sel and "매 입 처" in view_df.columns and not manufacturer_only:
        view_df = view_df[view_df["매 입 처"].isin(supplier_sel)]

    # 화면 표시 컬럼
    base_cols = ["제 조 사", "매 입 처", "상 품 명", "포장단위",
                 "재고수량", f"최근{selected_days}일_판매량", "기준판매량",
                 "최근입고수량", "부족수량", "과재고", "발주수량"]
    col_order = [c for c in base_cols if c in view_df.columns]
    # 미리보기에서도 매입처 숨기기 옵션 적용
    if manufacturer_only and "매 입 처" in col_order:
        col_order.remove("매 입 처")

    view_df = view_df.drop_duplicates(subset=["상 품 명", "포장단위"]).copy()
    sort_cols = [c for c in ["제 조 사", "매 입 처", "상 품 명"] if c in view_df.columns]
    if manufacturer_only and "매 입 처" in sort_cols:
        sort_cols.remove("매 입 처")
    if sort_cols:
        view_df = view_df[col_order].sort_values(sort_cols)
    else:
        view_df = view_df[col_order]

    # ===== 표 스타일 =====
    def style_df(df):
        def hi_short(v):
            try:
                v = int(v); return "background-color: #ffe5e5; font-weight: 700;" if v > 0 else ""
            except: return ""
        def hi_over(v):
            try:
                v = int(v); return "background-color: #eaf4ff;" if v > 0 else ""
            except: return ""
        num_cols = [c for c in [f"최근{selected_days}일_판매량","재고수량","기준판매량","최근입고수량","부족수량","과재고","발주수량"] if c in df.columns]
        styler = df.style.format("{:,}", subset=num_cols)
        if "부족수량" in df.columns:
            styler = styler.applymap(hi_short, subset=["부족수량"] + (["발주수량"] if "발주수량" in df.columns else []))
        if "과재고" in df.columns:
            styler = styler.applymap(hi_over, subset=["과재고"])
        return styler

    st.subheader("📊 발주 대상 미리보기")
    st.dataframe(style_df(view_df), use_container_width=True, height=520)

    # ===== 엑셀 내보내기 =====
    st.divider()
    st.subheader("📥 발주서 내보내기")
    if manufacturer_only:
        st.caption("양식에서 **매입처 열을 제외**하고, **제조사별**로만 파일/탭을 생성합니다.")
    else:
        st.caption("미리보기 기준대로 내보내지만, 권장 양식은 제조사 기준입니다.")

    # 실제 그룹 키와 내보내기 컬럼
    export_df = merged.copy()
    export_cols = [c for c in base_cols if c in export_df.columns]
    if manufacturer_only and "매 입 처" in export_cols:
        export_cols.remove("매 입 처")
    export_df = export_df[export_cols]

    # 그룹키 강제: 제조사
    if manufacturer_only:
        group_key = ["제 조 사"] if "제 조 사" in export_df.columns else ["상 품 명"]
    else:
        # 화면 옵션 따르되, 실제 존재하는 컬럼만
        if group_mode == "제조사":
            group_key = ["제 조 사"] if "제 조 사" in export_df.columns else ["상 품 명"]
        elif group_mode == "매입처" and "매 입 처" in export_df.columns:
            group_key = ["매 입 처"]
        else:
            if {"제 조 사","매 입 처"}.issubset(set(export_df.columns)):
                group_key = ["제 조 사","매 입 처"]
            else:
                group_key = ["제 조 사"] if "제 조 사" in export_df.columns else ["상 품 명"]

    if export_mode == "그룹별 개별 파일 (ZIP)":
        if st.button("ZIP 만들기"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for key_vals, group in export_df.groupby(group_key, dropna=False):
                    if not isinstance(key_vals, tuple):
                        key_vals = (key_vals,)
                    title = " - ".join([str(k) if (k is not None and str(k).strip()!='') else "미지정" for k in key_vals])
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                        sheet_df = group.copy()
                        # 매입처 컬럼 강제 제거(양식 일관성)
                        if "매 입 처" in sheet_df.columns:
                            sheet_df = sheet_df.drop(columns=["매 입 처"])
                        write_formatted_sheet(writer, "발주서", sheet_df)
                    safe_key = title.replace("/", "-")
                    filename = f"{safe_key} 발주서(최근{selected_days}일).xlsx"
                    zipf.writestr(filename, output.getvalue())
            zip_buffer.seek(0)
            st.download_button("📦 ZIP 파일 다운로드", data=zip_buffer, file_name=f"발주서_전체_최근{selected_days}일.zip", mime="application/zip")
    else:
        if st.button("엑셀(한 파일, 탭 구분) 만들기"):
            xls_buffer = io.BytesIO()
            with pd.ExcelWriter(xls_buffer, engine="xlsxwriter") as writer:
                for key_vals, group in export_df.groupby(group_key, dropna=False):
                    if not isinstance(key_vals, tuple):
                        key_vals = (key_vals,)
                    title = " - ".join([str(k) if (k is not None and str(k).strip()!='') else "미지정" for k in key_vals])
                    sheet_name = sanitize_sheet_name(title)
                    sheet_df = group.copy()
                    if "매 입 처" in sheet_df.columns:
                        sheet_df = sheet_df.drop(columns=["매 입 처"])
                    write_formatted_sheet(writer, sheet_name, sheet_df)
            xls_buffer.seek(0)
            st.download_button("📄 엑셀 다운로드", data=xls_buffer, file_name=f"발주서_탭구분_최근{selected_days}일.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ===== 간단 검증 요약 표시 =====
    with st.expander("🧪 데이터 검증 요약", expanded=False):
        st.dataframe(validate_and_report(sales_df, purchase_df, stock_df), use_container_width=True)
else:
    st.info(
        "📂 **좌측 사이드바**에서 **매출 자료, 매입 자료, 현재고** 파일을 모두 업로드하세요.\n\n"
        "매입처 컬럼이 없어도 **제조사 기준 발주서**를 생성할 수 있습니다."
    )
