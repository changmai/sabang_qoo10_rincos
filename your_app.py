import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="DR 자동 생성기", layout="wide")
st.title("📦 DR.XLSX 자동 생성 프로그램")

# 불필요한 텍스트 제거 함수
def clean_text(text):
    if not isinstance(text, str):
        return ""
    patterns = [
        r'#.*?セット', r'【.*?】', r'/.*?', r'韓コスメ', r'口紅', r'リップ', r'アワグロウ',
        r'[\[\]【】#]', r'\s{2,}'
    ]
    for pattern in patterns:
        text = re.sub(pattern, '', text)
    return re.sub(r'\s+', '', text.strip())

# 유사도 기반 매핑 함수
def match_items(source_list, target_list):
    mapping = {}
    for i, src in enumerate(source_list):
        best_score = 0
        best_match = None
        for j, tgt in enumerate(target_list):
            score = SequenceMatcher(None, src, tgt).ratio()
            if score > 0.3 and score > best_score:
                best_score = score
                best_match = j
        if best_match is not None:
            mapping[i] = best_match
    return mapping

# 기본 내재화된 H 데이터
@st.cache_data
def load_default_h():
    return pd.read_excel("H.xlsx")

# H 데이터 로딩
st.sidebar.header("H.XLSX 관리")
use_default_h = st.sidebar.checkbox("기본 내재화된 H.XLSX 사용", value=True)

if use_default_h:
    df_H = load_default_h()
else:
    h_file = st.sidebar.file_uploader("H.XLSX 파일 업로드", type=["xlsx"])
    if h_file:
        df_H = pd.read_excel(h_file)
    else:
        st.warning("H.XLSX 파일을 업로드하거나 기본 파일을 사용해주세요.")
        st.stop()

# S 파일 업로드
st.subheader("1단계: S.XLSX 파일 업로드")
s_file = st.file_uploader("S.XLSX 파일을 업로드하세요", type=["xlsx"])
if s_file:
    df_S = pd.read_excel(s_file)
    df_S.columns = df_S.columns.str.lower()
    df_H.columns = df_H.columns.str.lower()

    # 전처리
    s_item_names_raw = df_S['item_name'].fillna('').tolist()
    s_item_names_clean = [clean_text(name) for name in s_item_names_raw]
    h_names_clean = [clean_text(name) for name in df_H['출고상품명'].fillna('')]

    # 매핑
    s_to_h_map = match_items(s_item_names_clean, h_names_clean)

    # S 업데이트
    df_S_updated = df_S.copy()
    for s_idx, h_idx in s_to_h_map.items():
        df_S_updated.at[s_idx, '상품 shoppingmall url'] = df_H.at[h_idx, '상품 shoppingmall url']
        df_S_updated.at[s_idx, 'unit_total price'] = df_H.at[h_idx, 'unit_total price']

    df_S_updated['order_no'] = df_S_updated['order_no'].astype(str).apply(lambda x: '86' + x if not x.startswith('86') else x)
    df_S_updated['service code'] = df_S_updated.apply(lambda row: '99' if row.dropna().shape[0] > 1 else row['service code'], axis=1)
    df_S_updated['consignee_국가코드'] = df_S_updated.apply(lambda row: 'JP' if row.dropna().shape[0] > 1 else row['consignee_국가코드'], axis=1)
    df_S_updated['consignee_address (en)_jp지역 현지어 기재'] = df_S_updated['consignee_address (en)_jp지역 현지어 기재'].apply(lambda x: re.sub(r'\[.*?\]', '', x) if isinstance(x, str) else x)
    df_S_updated['pkg'] = df_S_updated.apply(lambda row: '1' if row.dropna().shape[0] > 1 else row['pkg'], axis=1)
    df_S_updated['item_origin'] = df_S_updated.apply(lambda row: 'KR' if row.dropna().shape[0] > 1 else row['item_origin'], axis=1)
    df_S_updated['currency'] = df_S_updated.apply(lambda row: 'JPY' if row.dropna().shape[0] > 1 else row['currency'], axis=1)

    # 우편번호 포맷 변경: 7자리 숫자 → XXX-XXXX
    def format_postal(postal):
        if isinstance(postal, (int, float)):
            postal = str(int(postal))
        if isinstance(postal, str) and postal.isdigit() and len(postal) == 7:
            return f"{postal[:3]}-{postal[3:]}"
        return postal

    df_S_updated['consignee_ postalcode'] = df_S_updated['consignee_ postalcode'].apply(format_postal)

    # DR 생성
    dr_columns = ['ref_no (주문번호)', '하이브 상품코드', '상품명', '수량', '바코드']
    df_DR = pd.DataFrame(columns=dr_columns)
    df_DR['ref_no (주문번호)'] = df_S_updated['order_no']
    df_DR['상품명'] = df_S_updated['item_name']
    df_DR['수량'] = df_S_updated['item_pcs']

    # DR 매핑 및 덮어쓰기
    dr_clean = [clean_text(str(x)) for x in df_DR['상품명'].fillna('')]
    dr_to_h_map = match_items(dr_clean, h_names_clean)
    for dr_idx, h_idx in dr_to_h_map.items():
        df_DR.at[dr_idx, '하이브 상품코드'] = df_H.at[h_idx, '상품코드']
        df_DR.at[dr_idx, '바코드'] = df_H.at[h_idx, '바코드']
        df_DR.at[dr_idx, '상품명'] = df_H.at[h_idx, '출고상품명']

    st.success("🎉 DR 파일 생성 완료!")
    st.dataframe(df_DR.head())

    today = datetime.today().strftime('%y%m%d')

    # S 파일 다운로드 버튼
    s_output = BytesIO()
    df_S_updated.to_excel(s_output, index=False)
    s_output.seek(0)
    st.download_button(
        label="📥 RINCOS_온드_주문등록양식_큐텐 다운로드",
        data=s_output,
        file_name=f"{today}_RINCOS_온드_주문등록양식_큐텐.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # DR 파일 다운로드 버튼
    dr_output = BytesIO()
    df_DR.to_excel(dr_output, index=False)
    dr_output.seek(0)
    st.download_button(
        label="📥 RINCOS_온드_HIVE센터 B2C 출고요청양식 다운로드",
        data=dr_output,
        file_name=f"{today}_RINCOS_온드_HIVE센터 B2C 출고요청양식.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("좌측에서 H.XLSX 설정 후, S.XLSX 파일을 업로드하세요.")
