
import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import timedelta

st.set_page_config(page_title="신명약품 자동발주(수량 중심)", layout="wide")

# ============= 헤더 =============
c1, c2 = st.columns([1, 5])
with c1:
    st.image("로고리뉴얼.png", width=100)
with c2:
    st.title("💊 신명약품 자동발주 – 수량 중심")
    st.caption("가격/단가 정보는 전부 제외하고, 현재고·매출수량·매입수량 대비 발주수량에만 집중합니다.")

st.sidebar.header("📂 파일 업로드")
sales_file = st.sidebar.file_uploader("매출자료 업로드", type=["xlsx"])
purchase_file = st.sidebar.file_uploader("매입자료 업로드", type=["xlsx"])
stock_file = st.sidebar.file_uploader("현재고 업로드", type=["xlsx"])

st.sidebar.divider()
mode = st.sidebar.radio("📅 분석 기간", ["자동 (최근 3개월)", "수동 지정"])

group_by_option = st.sidebar.radio("📁 발주서 그룹 기준", ["매 입 처", "제 조 사"])

st.sidebar.divider()
use_recent_purchase = st.sidebar.checkbox("최근 입고수량 반영하여 과발주 방지", value=True)
recent_days = st.sidebar.number_input("최근 입고 반영 일수", min_value=1, max_value=90, value=14, step=1)

st.sidebar.divider()
# =========================
# 변경: 발주 기준 선택을 1일~1년(365일) 범위 드롭다운으로 제공
# 기준판매량 = 최근 N일 매출수량 합계
# =========================
days_options = list(range(1, 366))
days_label_map = {d: f"{d}일" for d in days_options}
selected_days = st.sidebar.selectbox("발주 기준(최근 N일 판매량)", options=days_options, format_func=lambda x: days_label_map[x], index=29)  # 기본 30일

min_shortage = st.sidebar.number_input("부족수량 하한(이상만 표시)", min_value=0, value=0, step=1)
show_only_to_order = st.sidebar.checkbox("발주 필요 항목만 보기(부족수량>0)", value=True)

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

# ======== 메인 로직 ========
if sales_file and purchase_file and stock_file:
    sales_df = pd.read_excel(sales_file)  # 매출
    purchase_df = pd.read_excel(purchase_file)  # 매입(입고)
    stock_df = pd.read_excel(stock_file)  # 현재고

    # 컬럼 정규화
    sales_df = normalize_columns(sales_df, {
        "거래일자": "명세일자", "일자": "명세일자", "매출처": "매 출 처",
        "상품명": "상 품 명", "포장 단위": "포장단위"
    })
    purchase_df = normalize_columns(purchase_df, {
        "입고일": "입고일자", "거래처": "매 입 처", "상품명": "상 품 명",
        "포장 단위": "포장단위", "제조사": "제 조 사"
    })
    stock_df = normalize_columns(stock_df, {
        "거래처": "매 입 처", "상품명": "상 품 명",
        "포장 단위": "포장단위", "제조사": "제 조 사", "재고": "재고수량"
    })

    # 필수 컬럼 체크
    require_columns(sales_df, ["명세일자", "매 출 처", "상 품 명", "포장단위", "수량"], "매출자료")
    require_columns(purchase_df, ["입고일자", "매 입 처", "상 품 명", "제 조 사", "수량"], "매입자료")
    require_columns(stock_df, ["매 입 처", "제 조 사", "상 품 명", "포장단위", "재고수량"], "현재고")

    # 문자열 정리
    for df in [sales_df, purchase_df, stock_df]:
        df["상 품 명"] = to_upper_strip(df["상 품 명"])
        df["포장단위"] = to_upper_strip(df["포장단위"])

    # 날짜형
    sales_df["명세일자"] = pd.to_datetime(sales_df["명세일자"], errors="coerce")
    purchase_df["입고일자"] = pd.to_datetime(purchase_df["입고일자"], errors="coerce")

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

    sales_period = sales_df[(sales_df["명세일자"] >= start_date) & (sales_df["명세일자"] <= end_date)]

    # ===== 발주 기준: 최근 N일 판매량(총합) =====
    max_sale_date = sales_df["명세일자"].max()
    nday_start = max_sale_date - pd.Timedelta(days=int(selected_days))
    nday_sales = sales_df[(sales_df["명세일자"] > nday_start) & (sales_df["명세일자"] <= max_sale_date)]
    nday_qty = nday_sales.groupby(["상 품 명", "포장단위"], as_index=False)["수량"].sum()
    nday_qty.rename(columns={"수량": f"최근{selected_days}일_판매량"}, inplace=True)
    nday_qty["기준판매량"] = nday_qty[f"최근{selected_days}일_판매량"].astype(int)

    # 현재고 병합
    merged = nday_qty.merge(stock_df[["상 품 명", "포장단위", "재고수량", "매 입 처", "제 조 사"]].drop_duplicates(),
                            on=["상 품 명", "포장단위"], how="left")

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

    merged["발주수량"] = merged["부족수량"]  # 기본 로직: 부족=발주

    # 보기 옵션 필터
    if min_shortage > 0:
        merged = merged[merged["부족수량"] >= int(min_shortage)]
    if show_only_to_order:
        merged = merged[merged["발주수량"] > 0]

    # 중복 제거 및 정렬
    merged = merged.drop_duplicates(subset=["상 품 명", "포장단위"])
    dynamic_cols = ["제 조 사", "매 입 처", "상 품 명", "포장단위",
                    "재고수량", f"최근{selected_days}일_판매량", "기준판매량",
                    "최근입고수량", "부족수량", "과재고", "발주수량"]
    col_order = [c for c in dynamic_cols if c in merged.columns]
    merged = merged[col_order].sort_values(["제 조 사", "매 입 처", "상 품 명"])

    # ===== 상단 KPI =====
    k1, k2, k3, k4 = st.columns(4)
    total_items = len(merged)
    to_order_items = (merged["발주수량"] > 0).sum() if "발주수량" in merged else 0
    total_shortage = int(merged["부족수량"].sum()) if "부족수량" in merged else 0
    total_over = int(merged["과재고"].sum()) if "과재고" in merged else 0

    k1.metric("총 품목수", f"{total_items:,}")
    k2.metric("발주 필요 품목수", f"{to_order_items:,}")
    k3.metric("부족수량 합계", f"{total_shortage:,}")
    k4.metric("과재고 합계", f"{total_over:,}")

    # ===== 필터(검색/드롭다운) =====
    f1, f2, f3 = st.columns([2, 1, 1])
    with f1:
        keyword = st.text_input("🔎 상품명 검색(대소문자 무시)", value="").strip().upper()
    with f2:
        manu_sel = st.multiselect("제조사 필터", sorted(merged["제 조 사"].dropna().unique().tolist()))
    with f3:
        supplier_sel = st.multiselect("매입처 필터", sorted(merged["매 입 처"].dropna().unique().tolist()))

    view_df = merged.copy()
    if keyword:
        view_df = view_df[view_df["상 품 명"].str.contains(keyword, na=False)]
    if manu_sel:
        view_df = view_df[view_df["제 조 사"].isin(manu_sel)]
    if supplier_sel:
        view_df = view_df[view_df["매 입 처"].isin(supplier_sel)]

    # ===== 표 스타일(가독성 강화) =====
    def style_df(df):
        def highlight_shortage(v):
            if pd.isna(v):
                return ""
            try:
                v = int(v)
            except Exception:
                return ""
            if v > 0:
                return "background-color: #ffe5e5; font-weight: 700;"
            return ""
        def highlight_over(v):
            if pd.isna(v):
                return ""
            try:
                v = int(v)
            except Exception:
                return ""
            if v > 0:
                return "background-color: #eaf4ff;"
            return ""

        numeric_cols = [c for c in [f"최근{selected_days}일_판매량","재고수량","기준판매량","최근입고수량","부족수량","과재고","발주수량"] if c in df.columns]
        styler = df.style.format("{:,}", subset=numeric_cols)
        if "부족수량" in df.columns:
            styler = styler.applymap(highlight_shortage, subset=["부족수량", "발주수량"] if "발주수량" in df.columns else ["부족수량"])
        if "과재고" in df.columns:
            styler = styler.applymap(highlight_over, subset=["과재고"])
        return styler

    st.subheader("📊 발주 대상 미리보기")
    st.dataframe(style_df(view_df), use_container_width=True, height=520)

    # ===== ZIP 내보내기 =====
    st.divider()
    st.subheader("📥 발주서 일괄 다운로드 (ZIP)")
    st.caption("선택한 그룹 기준으로 시트를 구성한 개별 Excel 파일을 ZIP으로 묶어 제공합니다. (가격/단가 열 없음)")
    if st.button("ZIP 만들기"):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            # 그룹핑
            for key, group in merged.groupby(group_by_option):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    sheet_df = group.copy()
                    sheet_df = sheet_df[[c for c in col_order if c in sheet_df.columns]]
                    sheet_df.to_excel(writer, index=False, sheet_name="발주서")
                    ws = writer.sheets["발주서"]
                    for i, col in enumerate(sheet_df.columns):
                        maxlen = max(sheet_df[col].astype(str).map(len).max(), len(col))
                        ws.set_column(i, i, min(maxlen + 2, 40))
                safe_key = str(key).replace("/", "-")
                filename = f"{safe_key} 발주서(최근{selected_days}일).xlsx"
                zipf.writestr(filename, output.getvalue())
        zip_buffer.seek(0)
        st.download_button("📦 ZIP 파일 다운로드", data=zip_buffer, file_name=f"발주서_전체_최근{selected_days}일.zip", mime="application/zip")
else:
    st.info("📂 사이드바에서 **매출자료, 매입자료, 현재고** 파일을 모두 업로드하세요.")
