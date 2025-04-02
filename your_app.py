import streamlit as st
import pandas as pd
import difflib
import numpy as np
import re

# 내장 데이터: H.xlsx
h_data = {
    "출고상품명": ["아워 글로우 립 11 멜로우", "아워 글로우 립 15 윈디아"],
    "상품코드": ["O00100009", "O00100026"],
    "바코드": ["8809738598740", "8809864768703"],
    "URL": ["https://www.qoo10.jp/g/982447114/", "https://www.qoo10.jp/g/982447114/"],
    "UNIT_TOTAL PRICE": [1954, 1954]
}
df_h = pd.DataFrame(h_data)

# 내장 데이터: DR.xlsx
init_dr = {
    "Ref_No (주문번호)": ["1045038359", "1045038359"],
    "하이브 상품코드": ["", ""],
    "상품명": ["", ""],
    "수량": ["", ""],
    "바코드": ["", ""]
}
df_dr = pd.DataFrame(init_dr)

def normalize(text):
    return str(text).lower().replace(" ", "") if pd.notna(text) else ""

def remove_brackets(text):
    if pd.isna(text):
        return text
    return re.sub(r'\[[^\]]*\]', '', str(text))

def process(df_s):
    df_s = df_s.copy()
    h_names = df_h["출고상품명"].apply(normalize)
    s_names = df_s["ITEM_NAME"].fillna("").apply(normalize)

    sim_matrix = np.zeros((len(h_names), len(s_names)))
    for i, h in enumerate(h_names):
        for j, s in enumerate(s_names):
            sim_matrix[i, j] = difflib.SequenceMatcher(None, h, s).ratio()

    match_indexes = [(i, j) for i in range(len(h_names)) for j in range(len(s_names)) if sim_matrix[i, j] >= 0.3]

    for h_idx, s_idx in match_indexes:
        df_s.at[s_idx, "상품 Shoppingmall URL"] = df_h.at[h_idx, "URL"]
        df_s.at[s_idx, "UNIT_TOTAL PRICE"] = df_h.at[h_idx, "UNIT_TOTAL PRICE"]

    df_s["Order_No"] = df_s["Order_No"].apply(lambda x: f"86{x}" if pd.notna(x) and not str(x).startswith("86") else x)
    df_s["Service code"] = df_s.apply(lambda row: "99" if row.drop("Service code").notna().any() else row["Service code"], axis=1)
    df_s["CONSIGNEE_국가코드"] = df_s.apply(lambda row: "JP" if row.drop("CONSIGNEE_국가코드").notna().any() else row["CONSIGNEE_국가코드"], axis=1)
    df_s["CONSIGNEE_ADDRESS (EN)_JP지역 현지어 기재"] = df_s["CONSIGNEE_ADDRESS (EN)_JP지역 현지어 기재"].apply(remove_brackets)
    df_s["ITEM_ORIGIN"] = df_s.apply(lambda row: "1" if row.drop("ITEM_ORIGIN").notna().any() else row["ITEM_ORIGIN"], axis=1)
    df_s["상품 브랜드명"] = df_s.apply(lambda row: "KR" if row.drop("상품 브랜드명").notna().any() else row["상품 브랜드명"], axis=1)
    df_s["통관고유부호"] = df_s.apply(lambda row: "JPY" if row.drop("통관고유부호").notna().any() else row["통관고유부호"], axis=1)

    df_dr_updated = df_dr.copy()
    df_dr_updated["Ref_No (주문번호)"] = df_s["Order_No"]
    df_dr_updated["상품명"] = df_s["ITEM_NAME"]
    df_dr_updated["수량"] = df_s["ITEM_PCS"]

    dr_names = df_dr_updated["상품명"].apply(normalize)
    h_names = df_h["출고상품명"].apply(normalize)

    for i, dr_name in enumerate(dr_names):
        best_match_idx = np.argmax([difflib.SequenceMatcher(None, dr_name, h).ratio() for h in h_names])
        best_ratio = difflib.SequenceMatcher(None, dr_name, h_names[best_match_idx]).ratio()
        if best_ratio >= 0.3:
            df_dr_updated.at[i, "하이브 상품코드"] = df_h.at[best_match_idx, "상품코드"]
            df_dr_updated.at[i, "바코드"] = df_h.at[best_match_idx, "바코드"]
            df_dr_updated.at[i, "상품명"] = df_h.at[best_match_idx, "출고상품명"]

    return df_s, df_dr_updated

st.title("📦 S.XLSX 자동 매핑 웹 툴")

uploaded_file = st.file_uploader("S.XLSX 파일을 업로드하세요", type="xlsx")

if uploaded_file:
    df_s = pd.read_excel(uploaded_file)
    df_s_result, df_dr_result = process(df_s)

    st.success("🎉 매핑 및 처리 완료!")

    st.subheader("📋 업데이트된 S 파일")
    st.dataframe(df_s_result)
    st.download_button("📥 Updated_S.xlsx 다운로드", df_s_result.to_excel(index=False, engine='openpyxl'), file_name="Updated_S_Result.xlsx")

    st.subheader("📋 업데이트된 DR 파일")
    st.dataframe(df_dr_result)
    st.download_button("📥 Updated_DR.xlsx 다운로드", df_dr_result.to_excel(index=False, engine='openpyxl'), file_name="Updated_DR_Result.xlsx")
