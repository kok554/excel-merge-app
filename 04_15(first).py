import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="비품 수량 자동 병합기", layout="wide")

st.title("🏢 참가업체 비품 주문서 자동 병합기 (최종 구조 대응)")

uploaded_files = st.file_uploader("비품 주문서 파일 업로드 (여러 개 선택)", type=["xlsx"], accept_multiple_files=True)

@st.cache_data
def process_file(file):
    try:
        df = pd.read_excel(file, sheet_name='1부스', header=None)

        company_name = df.iloc[7, 1] if not pd.isna(df.iloc[7, 1]) else "업체명 미기재"

        # 실제 데이터는 17행부터 시작 (인덱스 16)
        temp_df = df.iloc[16:36, [0, 2, 4]].copy()
        temp_df.columns = ['품목', '기본제공수량', '추가요청수량']
        temp_df = temp_df.dropna(subset=['품목'])

        def extract_sum(x):
            if isinstance(x, str):
                nums = re.findall(r'\d+', x)
                return sum(map(int, nums)) if nums else 0
            if pd.isna(x):
                return 0
            return int(x)

        temp_df['기본제공수량'] = temp_df['기본제공수량'].apply(extract_sum)
        temp_df['추가요청수량'] = temp_df['추가요청수량'].apply(extract_sum)
        temp_df['총수량'] = temp_df['기본제공수량'] + temp_df['추가요청수량']
        temp_df['업체명'] = company_name

        return temp_df[temp_df['총수량'] > 0]
    except Exception:
        return None

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        result = process_file(file)
        if result is not None:
            all_data.append(result)
        else:
            st.warning(f"{file.name} 파일 처리 중 문제 발생, 건너뜀.")

    if all_data:
        result_df = pd.concat(all_data, ignore_index=True)

        st.subheader("🏷️ 업체별 비품 수량 보기")

        companies = result_df['업체명'].unique().tolist()
        selected_companies = st.multiselect("업체를 선택하세요", companies, default=companies)

        for company in selected_companies:
            with st.expander(f"🏢 {company}", expanded=False):
                st.dataframe(result_df[result_df['업체명'] == company][['품목', '기본제공수량', '추가요청수량', '총수량']])

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Summary')
            processed_data = output.getvalue()
            return processed_data

        st.download_button(
            label="📥 전체 병합된 엑셀 다운로드",
            data=to_excel(result_df),
            file_name="업체별_비품_수량_통합.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.error("처리 가능한 파일이 없습니다.")
else:
    st.info("좌측 사이드바에서 파일을 업로드하세요.")
