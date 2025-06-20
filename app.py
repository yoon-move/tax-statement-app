import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
from datetime import datetime

st.set_page_config(page_title="세금계산서 & 통장 매칭", layout="wide", initial_sidebar_state="expanded")

st.title("📊 세금계산서 & 통장 거래내역 자동 정산")
st.markdown("업로드된 세금계산서와 거래내역 데이터를 자동으로 정산합니다.")

# --- 파일 업로드 ---
st.sidebar.header("📂 파일 업로드")
sell_file = st.sidebar.file_uploader("💼 매출 세금계산서 (.xlsx)", type=["xlsx"])
buy_file = st.sidebar.file_uploader("🧾 매입 세금계산서 (.xlsx)", type=["xlsx"])
bank_biz_file = st.sidebar.file_uploader("🏦 사업자 통장 거래내역 (.xlsx)", type=["xls", "xlsx"])
bank_tg_file = st.sidebar.file_uploader("🏛️ 기보 통장 거래내역 (.xlsx)", type=["xls", "xlsx"])
uploaded = st.button("📤 업로드 완료", type="primary")

# --- 파일 로딩 함수 ---
def load_invoice(file, label):
    try:
        df = pd.read_excel(file, skiprows=6)
        df = df[df.columns[df.columns.str.contains("작성일자") | df.columns.str.contains("상호") | df.columns.str.contains("합계금액")]]
        df.columns = df.columns.str.strip()
        df = df.rename(columns=lambda x: x.replace(" ", ""))
        df["작성일자"] = pd.to_datetime(df["작성일자"], errors="coerce").dt.date
        df["합계금액"] = pd.to_numeric(df["합계금액"], errors="coerce")
        df["구분"] = label
        df = df.rename(columns={col: "거래처명" for col in df.columns if "상호" in col})
        return df[["작성일자", "거래처명", "합계금액", "구분"]]
    except Exception as e:
        st.warning(f"{label} 세금계산서 불러오기 오류: {e}")
        return pd.DataFrame()

def load_bank(file, label):
    try:
        df = pd.read_excel(file, skiprows=6)
        df.columns = df.columns.str.strip()
        df = df.rename(columns={"보낸분/받는분": "거래처명"})
        df["거래일자"] = pd.to_datetime(df["거래일시"], errors="coerce").dt.date
        df["입금액"] = pd.to_numeric(df.get("입금액(원)", 0), errors="coerce").fillna(0)
        df["출금액"] = pd.to_numeric(df.get("출금액(원)", 0), errors="coerce").fillna(0)
        df["거래금액"] = df["입금액"] - df["출금액"]
        df["계좌구분"] = label
        return df[["거래일자", "거래처명", "거래금액", "계좌구분"]]
    except Exception as e:
        st.warning(f"{label} 통장 불러오기 오류: {e}")
        return pd.DataFrame()

# --- 정규화 함수 ---
def normalize(name):
    if not isinstance(name, str):
        return ""
    name = name.strip().replace("(주)", "").replace("주식회사", "").replace("농업회사법인", "").replace("종합상사", "")
    for exc in ["네이버", "네이버파이낸셜"]:
        if name.strip() == exc:
            return exc
    return name.replace(" ", "").lower()

# --- 실행 ---
if uploaded and (sell_file or buy_file) and (bank_biz_file or bank_tg_file):
    invoice_df = pd.concat([
        load_invoice(sell_file, "매출") if sell_file else pd.DataFrame(),
        load_invoice(buy_file, "매입") if buy_file else pd.DataFrame()
    ], ignore_index=True)

    bank_df = pd.concat([
        load_bank(bank_biz_file, "사업자통장") if bank_biz_file else pd.DataFrame(),
        load_bank(bank_tg_file, "기보통장") if bank_tg_file else pd.DataFrame()
    ], ignore_index=True)

    invoice_df["정규화거래처명"] = invoice_df["거래처명"].apply(normalize)
    bank_df["정규화거래처명"] = bank_df["거래처명"].apply(normalize)

    # ✅ 집계
    inv_sum = invoice_df.groupby("정규화거래처명")["합계금액"].sum().reset_index()
    inv_sum.columns = ["정규화거래처명", "세금계산서_합계"]

    bank_sum = bank_df.groupby("정규화거래처명")["거래금액"].sum().reset_index()
    bank_sum.columns = ["정규화거래처명", "거래내역_합계"]

    summary = pd.merge(inv_sum, bank_sum, on="정규화거래처명", how="inner")
    summary["차이"] = summary["세금계산서_합계"] - summary["거래내역_합계"]
    summary["정산결과"] = summary["차이"].apply(lambda x: "✅ 일치" if abs(x) < 1000 else "❌ 미일치")

    summary["세금계산서_합계"] = summary["세금계산서_합계"].map("{:,.0f}원".format)
    summary["거래내역_합계"] = summary["거래내역_합계"].map("{:,.0f}원".format)
    summary["차이"] = summary["차이"].map("{:,.0f}원".format)

    st.subheader("📑 거래처별 정산 결과")
    st.dataframe(summary, use_container_width=True)

    # -----------------------------
    # 🛠️ 수동 거래처 매칭 보정 기능
    # -----------------------------

    st.subheader("🛠️ 수동 거래처 매칭 보정")

    unmatched_invoices = invoice_df["거래처명"].dropna().unique().tolist()
    unmatched_banks = bank_df["거래처명"].dropna().unique().tolist()

    selected_invoice = st.selectbox("📤 세금계산서 거래처명 선택", [""] + unmatched_invoices, key="invoice_select")
    selected_bank = st.selectbox("🏦 통장 거래처명 선택", [""] + unmatched_banks, key="bank_select")

    if st.button("✅ 수동 매칭 적용", key="apply_manual_match") and selected_invoice and selected_bank:
        corrected_name = f"수정:{selected_invoice.strip()}=={selected_bank.strip()}"
        invoice_df.loc[invoice_df["거래처명"] == selected_invoice, "정규화거래처명"] = corrected_name
        bank_df.loc[bank_df["거래처명"] == selected_bank, "정규화거래처명"] = corrected_name

        st.success(f"✅ '{selected_invoice}' 와(과) '{selected_bank}' 을(를) 수동으로 연결했습니다.")

        inv_sum = invoice_df.groupby("정규화거래처명")["합계금액"].sum().reset_index()
        inv_sum.columns = ["정규화거래처명", "세금계산서_합계"]

        bank_sum = bank_df.groupby("정규화거래처명")["거래금액"].sum().reset_index()
        bank_sum.columns = ["정규화거래처명", "거래내역_합계"]

        corrected_summary = pd.merge(inv_sum, bank_sum, on="정규화거래처명", how="inner")
        corrected_summary["차이"] = corrected_summary["세금계산서_합계"] - corrected_summary["거래내역_합계"]
        corrected_summary["정산결과"] = corrected_summary["차이"].apply(lambda x: "✅ 일치" if abs(x) < 1000 else "❌ 미일치")

        corrected_summary["세금계산서_합계"] = corrected_summary["세금계산서_합계"].map("{:,.0f}원".format)
        corrected_summary["거래내역_합계"] = corrected_summary["거래내역_합계"].map("{:,.0f}원".format)
        corrected_summary["차이"] = corrected_summary["차이"].map("{:,.0f}원".format)

        st.markdown("### 🔁 수동 보정 반영된 정산 결과")
        st.dataframe(corrected_summary, use_container_width=True)
