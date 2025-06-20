import streamlit as st
import pandas as pd
import numpy as np
from io import StringIO
from datetime import datetime
import plotly.express as px
from fpdf import FPDF
import tempfile
import os
import io

st.set_page_config(page_title="세금계산서 & 은행거래 비교", layout="wide")

st.title("📊 세금계산서 & 은행 계좌 내역 통합관리")
st.markdown("세금계산서 CSV와 은행 거래내역 CSV 파일을 업로드하여 거래 일치 여부를 분석합니다.")

# --- 파일 업로드 ---
st.sidebar.header("📂 파일 업로드")
sell_file = st.sidebar.file_uploader("매출 세금계산서 XLSX 업로드", type=["xlsx"])
buy_file = st.sidebar.file_uploader("매입 세금계산서 XLSX 업로드", type=["xlsx"])
bank_biz_file = st.sidebar.file_uploader("사업자통장 거래내역 XLS 또는 CSV 업로드", type=["xls", "xlsx", "csv"])
bank_tg_file = st.sidebar.file_uploader("기보통장 거래내역 XLS 또는 CSV 업로드", type=["xls", "xlsx", "csv"])

def load_invoice_data(file, label):
    try:
        xl = pd.ExcelFile(file)
        for i in range(5, 20):
            df = pd.read_excel(xl, sheet_name="세금계산서", header=i)
            if "작성일자" in df.columns and "공급가액" in df.columns:
                if "상호.1" in df.columns:
                    df = df[["작성일자", "공급자사업자등록번호", "상호", "대표자명", "공급받는자사업자등록번호", "상호.1", "공급가액", "세액", "합계금액"]].copy()
                    df.columns = ["작성일자", "공급자사업자등록번호", "공급자 상호", "공급자 대표자명", "공급받는자사업자등록번호", "공급받는자 상호", "공급가액", "세액", "합계금액"]
                    df["구분"] = label
                    return df
    except Exception as e:
        st.warning(f"{label} 세금계산서 불러오기 실패: {e}")
    return pd.DataFrame()

def load_bank_file(file, label):
    try:
        if file.name.endswith(".csv"):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        df['계좌구분'] = label
        df['거래일자'] = pd.to_datetime(df['거래일자'], errors='coerce')
        return df
    except Exception as e:
        st.error(f"{label} 통장을 불러오는 중 오류 발생: {e}")
        return pd.DataFrame()

if sell_file and buy_file and bank_biz_file and bank_tg_file:
    sell_df = load_invoice_data(sell_file, "매출")
    buy_df = load_invoice_data(buy_file, "매입")
    invoice_df = pd.concat([sell_df, buy_df], ignore_index=True)

    # --- 내부거래 필터링 ---
    mask = (
        invoice_df["공급자사업자등록번호"].astype(str).str.contains("447-87-03172", na=False) |
        invoice_df["공급받는자사업자등록번호"].astype(str).str.contains("447-87-03172", na=False) |
        invoice_df["공급자 상호"].astype(str).str.contains("그로와이즈", na=False) |
        invoice_df["공급받는자 상호"].astype(str).str.contains("그로와이즈", na=False) |
        invoice_df["공급자 대표자명"].astype(str).str.contains("윤영범", na=False)
    )
    invoice_df = invoice_df[~mask].copy()

    invoice_df['작성일자'] = pd.to_datetime(invoice_df['작성일자'], errors='coerce')

    # --- 은행 파일 통합 로딩 ---
    bank_biz_df = load_bank_file(bank_biz_file, "사업자통장")
    bank_tg_df = load_bank_file(bank_tg_file, "기보통장")
    bank_df = pd.concat([bank_biz_df, bank_tg_df], ignore_index=True)

    # --- 일치 여부 판별 ---
    def match_rows(inv, bank):
        results = []
        for i, row in inv.iterrows():
            matched = bank[
                (bank['거래처명'] == row['공급받는자 상호']) &
                (np.abs((bank['거래일자'] - row['작성일자']).dt.days) <= 1) &
                (bank['입금액'] == row['합계금액'])
            ]
            if not matched.empty:
                results.append("✅ 일치")
            else:
                partial = bank[
                    (bank['거래처명'] == row['공급받는자 상호']) &
                    (np.abs((bank['거래일자'] - row['작성일자']).dt.days) <= 3)
                ]
                if not partial.empty:
                    results.append("⚠️ 일부일치")
                else:
                    results.append("❌ 미일치")
        return results

    invoice_df['매칭결과'] = match_rows(invoice_df, bank_df)

    # --- 필터 ---
    st.sidebar.header("🔍 검색 필터")
    filter_match = st.sidebar.multiselect("매칭 결과 필터", options=invoice_df['매칭결과'].unique(), default=invoice_df['매칭결과'].unique())
    filter_vendor = st.sidebar.text_input("거래처명 검색")
    filtered_df = invoice_df[invoice_df['매칭결과'].isin(filter_match)]
    if filter_vendor:
        filtered_df = filtered_df[filtered_df['공급받는자 상호'].str.contains(filter_vendor, case=False, na=False)]

    # --- 출력 ---
    st.subheader("📑 세금계산서 매칭 결과")
    st.dataframe(filtered_df, use_container_width=True)

    st.markdown("### 📌 매칭 통계")
    st.write(filtered_df['매칭결과'].value_counts())

    st.markdown("### 📈 월별 매출 추이")
    filtered_df['월'] = filtered_df['작성일자'].dt.to_period('M').astype(str)
    monthly_sum = filtered_df.groupby('월')['합계금액'].sum().reset_index()
    fig = px.bar(monthly_sum, x='월', y='합계금액', text='합계금액', title='월별 세금계산서 합계')
    st.plotly_chart(fig, use_container_width=True)

    # --- 다운로드 ---
    csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button("📥 결과 CSV 다운로드", data=csv, file_name="매칭결과.csv", mime="text/csv")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='매칭결과')
    st.download_button(
        label="📥 결과 Excel 다운로드",
        data=output.getvalue(),
        file_name="매칭결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("왼쪽 사이드바에서 매입, 매출, 그리고 두 종류의 은행 거래내역 파일을 모두 업로드해주세요.")
