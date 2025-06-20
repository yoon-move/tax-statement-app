import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import io

st.set_page_config(page_title="세금계산서 & 은행거래 비교", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stFileUploader > label div:first-child {
        background-color: #fff3e0;
        border: 1px dashed #ff9800;
        padding: 12px;
        transition: background-color 0.3s;
        color: black !important;
        font-weight: 500;
    }
    .stFileUploader > label div:first-child:hover {
        background-color: #ffe0b2 !important;
    }
    .stFileUploader .uploadedFileName {
        color: black !important;
    }
    .stFileUploader input[type="file"]::file-selector-button {
        color: black;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 세금계산서 & 은행 계좌 내역 통합관리")
st.markdown("세금계산서와 은행 거래내역 파일을 업로드하여 거래 일치 여부를 분석합니다.")

# --- 파일 업로드 영역 ---
st.sidebar.header("📂 파일 업로드")
sell_file = st.sidebar.file_uploader("💼 매출 세금계산서 업로드 (엑셀 파일 .xlsx)", type=["xlsx"])
buy_file = st.sidebar.file_uploader("🧾 매입 세금계산서 업로드 (엑셀 파일 .xlsx)", type=["xlsx"])
bank_biz_file = st.sidebar.file_uploader("🏦 사업자 통장 거래내역 업로드 (.xlsx)", type=["xls", "xlsx"])
bank_tg_file = st.sidebar.file_uploader("🏛️ 기보 통장 거래내역 업로드 (.xlsx)", type=["xls", "xlsx"])

uploaded = st.button("📤 업로드 완료", type="primary")

# 이름 정규화 함수
def normalize(name):
    if pd.isna(name):
        return ""
    return (
        str(name)
        .lower()
        .replace("(주)", "")
        .replace("주식회사", "")
        .replace("(", "")
        .replace(")", "")
        .replace(" ", "")
        .strip()
    )

# 세금계산서 불러오기
def load_invoice(file, label):
    try:
        df = pd.read_excel(file, sheet_name="세금계산서", header=5)
        df = df[df["합계금액"] > 0].copy()
        df["구분"] = label
        df["작성일자"] = pd.to_datetime(df["작성일자"], errors="coerce")
        df["정규이름"] = df["상호"].apply(normalize)
        df["대표자명"] = df["대표자명"].fillna("")
        return df
    except Exception as e:
        st.error(f"{label} 세금계산서 불러오기 실패: {e}")
        return pd.DataFrame()

# 통장 불러오기
def load_bank(file, label):
    try:
        df = pd.read_excel(file, skiprows=6)
        df = df.rename(columns={
            "거래일시": "거래일자",
            "보낸분/받는분": "거래처명",
            "입금액(원)": "입금액",
            "출금액(원)": "출금액"
        })
        df["거래일자"] = pd.to_datetime(df["거래일자"], errors="coerce")
        df["입금액"] = pd.to_numeric(df["입금액"], errors="coerce")
        df["출금액"] = pd.to_numeric(df["출금액"], errors="coerce")
        df["거래금액"] = df["입금액"].fillna(0) - df["출금액"].fillna(0)
        df["정규이름"] = df["거래처명"].apply(normalize)
        df["계좌구분"] = label
        return df[["거래일자", "거래처명", "거래금액", "정규이름", "계좌구분"]]
    except Exception as e:
        st.error(f"{label} 통장 불러오기 실패: {e}")
        return pd.DataFrame()

# --- 본처리 ---
if uploaded and ((sell_file or buy_file) and (bank_biz_file or bank_tg_file)):

    # 세금계산서 로딩
    df_invoice = pd.DataFrame()
    if sell_file:
        df_invoice = pd.concat([df_invoice, load_invoice(sell_file, "매출")], ignore_index=True)
    if buy_file:
        df_invoice = pd.concat([df_invoice, load_invoice(buy_file, "매입")], ignore_index=True)

    # 통장 로딩
    df_bank = pd.DataFrame()
    if bank_biz_file:
        df_bank = pd.concat([df_bank, load_bank(bank_biz_file, "사업자통장")], ignore_index=True)
    if bank_tg_file:
        df_bank = pd.concat([df_bank, load_bank(bank_tg_file, "기보통장")], ignore_index=True)

    # 매칭 수행
    match_results = []
    for _, row in df_invoice.iterrows():
        date_range = pd.date_range(row["작성일자"] - pd.Timedelta(days=1), row["작성일자"] + pd.Timedelta(days=1))
        candidates = df_bank[(df_bank["정규이름"] == row["정규이름"]) & (df_bank["거래일자"].isin(date_range))]

        if not candidates[candidates["거래금액"] == row["합계금액"]].empty:
            match_results.append("✅ 일치")
        elif not candidates.empty:
            match_results.append("⚠️ 일부일치")
        else:
            match_results.append("❌ 미일치")

    df_invoice["매칭결과"] = match_results

    # 결과 출력
    st.subheader("📑 매칭 결과")
    st.dataframe(df_invoice[["작성일자", "상호", "대표자명", "합계금액", "구분", "매칭결과"]], use_container_width=True)

    # 통계
    st.markdown("### 📈 월별 매출 추이")
    df_invoice["월"] = df_invoice["작성일자"].dt.to_period("M").astype(str)
    monthly_sum = df_invoice.groupby("월")["합계금액"].sum().reset_index()
    fig = px.bar(monthly_sum, x="월", y="합계금액", text="합계금액", title="월별 세금계산서 합계")
    st.plotly_chart(fig, use_container_width=True)

    # 다운로드
    csv = df_invoice.to_csv(index=False).encode("utf-8-sig")
    st.download_button("📥 결과 CSV 다운로드", data=csv, file_name="매칭결과.csv", mime="text/csv")

else:
    st.info("왼쪽 사이드바에서 세금계산서와 통장 거래내역 중 최소 1개씩 업로드한 후 '📤 업로드 완료' 버튼을 눌러주세요.")
