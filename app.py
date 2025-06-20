# app.py

import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="세금계산서 & 계좌 비교", layout="wide", initial_sidebar_state="expanded")

# 👉 스타일
st.markdown("""
<style>
.stFileUploader > label div:first-child {
    background-color: #fff3e0;
    border: 1px dashed #ff9800;
    padding: 12px;
    color: black;
    font-weight: 500;
}
.stFileUploader > label div:first-child:hover {
    background-color: #ffe0b2 !important;
}
</style>
""", unsafe_allow_html=True)

st.title("📊 세금계산서 & 은행 거래내역 통합 비교")
st.markdown("세금계산서와 거래내역을 비교하여 거래처별로 매칭 결과를 확인하세요.")

# 👉 파일 업로드
st.sidebar.header("📂 파일 업로드")
sell_file = st.sidebar.file_uploader("💼 매출 세금계산서 (.xlsx)", type=["xlsx"])
buy_file = st.sidebar.file_uploader("🧾 매입 세금계산서 (.xlsx)", type=["xlsx"])
bank1 = st.sidebar.file_uploader("🏦 사업자 통장 거래내역 (.xlsx)", type=["xlsx"])
bank2 = st.sidebar.file_uploader("🏛️ 기보 통장 거래내역 (.xlsx)", type=["xlsx"])
uploaded = st.button("📤 업로드 완료", type="primary")

# 👉 헤더 탐지 함수
def find_invoice_header_row(file, required_cols=None, max_rows=20):
    if required_cols is None:
        required_cols = ["작성일자", "공급가액", "합계금액"]
    xl = pd.ExcelFile(file)
    for i in range(max_rows):
        try:
            df = pd.read_excel(xl, sheet_name=0, header=i, nrows=5)
            if all(col in df.columns for col in required_cols):
                return i
        except:
            continue
    return None

# 👉 컬럼 정규화
def rename_invoice_columns(df):
    if "상호.1" in df.columns:
        df["공급받는자 상호"] = df["상호.1"]
    if "공급받는자사업자등록번호" not in df.columns:
        df["공급받는자사업자등록번호"] = ""
    return df

# 👉 세금계산서 불러오기
def load_invoice(file):
    header_row = find_invoice_header_row(file)
    if header_row is None:
        return pd.DataFrame()
    df = pd.read_excel(file, header=header_row)
    df = rename_invoice_columns(df)
    df["작성일자"] = pd.to_datetime(df["작성일자"], errors="coerce")
    df["합계금액"] = pd.to_numeric(df["합계금액"], errors="coerce")
    return df

# 👉 거래내역 불러오기
def load_bank(file, label):
    df = pd.read_excel(file, skiprows=6)
    df.columns = df.columns.str.strip()
    # 자동 컬럼 매핑
    if "거래일자" not in df.columns and "거래일시" in df.columns:
        df["거래일자"] = df["거래일시"]
    if "보낸분/받는분" in df.columns:
        df["거래처명"] = df["보낸분/받는분"]
    elif "받는분" in df.columns:
        df["거래처명"] = df["받는분"]
    if "입금액(원)" in df.columns:
        df["입금액"] = pd.to_numeric(df["입금액(원)"], errors="coerce")
    if "출금액(원)" in df.columns:
        df["출금액"] = pd.to_numeric(df["출금액(원)"], errors="coerce")

    df["거래일자"] = pd.to_datetime(df["거래일자"], errors="coerce")
    df["계좌"] = label
    return df

# 👉 거래처 필터링 보정 (유사도 규칙)
def normalize_vendor_name(name):
    ignore_words = ["주식회사", "(주)", "농업회사법인", "종합상사"]
    for word in ignore_words:
        name = name.replace(word, "")
    return name.strip()

# 👉 거래처별 매칭
def match_by_vendor(invoice_df, bank_df, invoice_type="매입"):
    # 매입이면 입금액, 매출이면 출금액
    amt_col = "입금액" if invoice_type == "매입" else "출금액"

    invoice_df = invoice_df.copy()
    bank_df = bank_df.copy()

    invoice_df["상호정규화"] = invoice_df["공급받는자 상호"].fillna("").apply(normalize_vendor_name)
    bank_df["상호정규화"] = bank_df["거래처명"].fillna("").apply(normalize_vendor_name)

    result = []
    for vendor in invoice_df["상호정규화"].unique():
        inv_sum = invoice_df[invoice_df["상호정규화"] == vendor]["합계금액"].sum()
        bank_sum = bank_df[bank_df["상호정규화"] == vendor][amt_col].sum()
        matched = abs(inv_sum - bank_sum) < 1000  # 오차 허용범위
        result.append({
            "거래처": vendor,
            "세금계산서합계": inv_sum,
            "통장거래합계": bank_sum,
            "일치여부": "✅ 일치" if matched else "❌ 불일치"
        })

    return pd.DataFrame(result)

# 👉 실행 로직
if uploaded and (sell_file or buy_file) and (bank1 or bank2):
    inv_df = pd.DataFrame()
    if buy_file:
        inv_df = load_invoice(buy_file)
    if sell_file:
        inv_df = pd.concat([inv_df, load_invoice(sell_file)], ignore_index=True)

    bank_df = pd.DataFrame()
    if bank1:
        bank_df = load_bank(bank1, "사업자")
    if bank2:
        bank_df = pd.concat([bank_df, load_bank(bank2, "기보")], ignore_index=True)

    if not inv_df.empty and not bank_df.empty:
        st.subheader("📊 거래처별 매칭 결과")
        report_df = match_by_vendor(inv_df, bank_df, invoice_type="매입")
        st.dataframe(report_df)

        csv = report_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("📥 결과 CSV 다운로드", data=csv, file_name="거래처_매칭_결과.csv", mime="text/csv")
    else:
        st.warning("불러온 데이터가 없습니다. 파일을 다시 확인해 주세요.")
else:
    st.info("왼쪽 사이드바에서 세금계산서와 통장 거래내역을 각각 1개 이상 업로드한 뒤 '업로드 완료'를 눌러주세요.")
