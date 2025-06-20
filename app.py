import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import io
from difflib import get_close_matches

st.set_page_config(page_title="세금계산서 & 은행거래 비교", layout="wide", initial_sidebar_state="expanded")

# ------------------------- 스타일 -------------------------
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

# ------------------------- 유틸 함수 -------------------------
def find_invoice_header_row(file, required_cols=None, max_rows=20):
    if required_cols is None:
        required_cols = ["작성일자", "공급가액", "합계금액"]
    try:
        xl = pd.ExcelFile(file)
        for i in range(max_rows):
            try:
                df = pd.read_excel(xl, header=i, nrows=5)
                if all(col in df.columns for col in required_cols):
                    return i
            except:
                continue
        return None
    except Exception as e:
        st.error(f"헤더 찾기 오류: {e}")
        return None

# ------------------------- 파일 업로드 -------------------------
st.sidebar.header("📂 파일 업로드")
sell_file = st.sidebar.file_uploader("💼 매출 세금계산서 업로드 (.xlsx)", type=["xlsx"])
buy_file = st.sidebar.file_uploader("🧾 매입 세금계산서 업로드 (.xlsx)", type=["xlsx"])
bank_biz_file = st.sidebar.file_uploader("🏦 사업자 통장 거래내역 (.xlsx)", type=["xlsx"])
bank_tg_file = st.sidebar.file_uploader("🏛️ 기보 통장 거래내역 (.xlsx)", type=["xlsx"])
uploaded = st.button("📤 업로드 완료", type="primary")

# ------------------------- 데이터 처리 -------------------------
def load_invoice(file, label):
    header_row = find_invoice_header_row(file)
    if header_row is None:
        st.warning(f"{label} 세금계산서에서 작성일자/공급가액/합계금액을 찾을 수 없습니다.")
        return pd.DataFrame()
    df = pd.read_excel(file, header=header_row)
    df["작성일자"] = pd.to_datetime(df["작성일자"], errors="coerce")
    df["작성일자"] = df["작성일자"].dt.strftime("%Y-%m-%d")
    df["구분"] = label
    return df

def load_bank(file, label):
    try:
        df = pd.read_excel(file, skiprows=6)
        df.columns = df.columns.str.strip()
        df = df.rename(columns=lambda x: x.replace("(원)", "").strip())
        df = df.rename(columns={"보낸분/받는분": "거래처명", "입금액": "입금", "출금액": "출금", "거래일시": "거래일자"})
        df["거래일자"] = pd.to_datetime(df["거래일자"], errors="coerce")
        df["계좌"] = label
        return df
    except Exception as e:
        st.error(f"{label} 통장 불러오기 오류: {e}")
        return pd.DataFrame()

def normalize_name(name):
    if not isinstance(name, str): return ""
    remove_words = ["주식회사", "(주)", "농업회사법인", "종합상사", "㈜"]
    for word in remove_words:
        name = name.replace(word, "")
    return name.strip()

def match_by_vendor(invoice_df, bank_df):
    result = []
    grouped = invoice_df.groupby(["공급받는자 상호", "공급받는자사업자등록번호"])
    for (vendor, code), group in grouped:
        total_invoice_amt = group["합계금액"].sum()
        target_name = normalize_name(vendor)

        match_names = bank_df["거래처명"].dropna().apply(normalize_name)
        matches = bank_df[match_names == target_name]

        total_in = matches["입금"].fillna(0).sum()
        total_out = matches["출금"].fillna(0).sum()
        total_bank_amt = total_in - total_out

        matched = np.isclose(total_invoice_amt, abs(total_bank_amt), atol=1000)
        result.append({
            "거래처명": vendor,
            "사업자번호": code,
            "세금계산서합계": total_invoice_amt,
            "통장거래합계": total_bank_amt,
            "매칭여부": "✅ 일치" if matched else "❌ 불일치"
        })
    return pd.DataFrame(result)

# ------------------------- 실행 -------------------------
if uploaded:
    invoice_df = pd.DataFrame()
    if sell_file: invoice_df = pd.concat([invoice_df, load_invoice(sell_file, "매출")])
    if buy_file: invoice_df = pd.concat([invoice_df, load_invoice(buy_file, "매입")])

    bank_df = pd.DataFrame()
    if bank_biz_file: bank_df = pd.concat([bank_df, load_bank(bank_biz_file, "사업자통장")])
    if bank_tg_file: bank_df = pd.concat([bank_df, load_bank(bank_tg_file, "기보통장")])

    if not invoice_df.empty and not bank_df.empty:
        report_df = match_by_vendor(invoice_df, bank_df)
        st.subheader("📋 거래처별 세금계산서 & 통장 매칭 리포트")
        st.dataframe(report_df, use_container_width=True)

        # ------------------------- 수동 매칭 보정 -------------------------
        st.markdown("### 🛠️ 수동 매칭 보정")
        selected = st.selectbox("매칭 오류로 보정할 거래처 선택", options=report_df["거래처명"].unique())
        if selected:
            inv_amt = report_df.loc[report_df["거래처명"] == selected, "세금계산서합계"].values[0]
            st.write(f"📄 세금계산서 금액: {inv_amt}")
            new_amt = st.number_input("거래내역에서 수동 금액 입력 (수동 매칭용)", value=0)
            if st.button("📌 수동 매칭 적용"):
                report_df.loc[report_df["거래처명"] == selected, "통장거래합계"] = new_amt
                report_df.loc[report_df["거래처명"] == selected, "매칭여부"] = (
                    "✅ 일치" if np.isclose(inv_amt, new_amt, atol=1000) else "❌ 불일치"
                )
                st.success("수동 매칭이 반영되었습니다.")
                st.dataframe(report_df, use_container_width=True)
    else:
        st.warning("파일이 부족하거나 잘못 업로드되었습니다.")
