import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="비품 수량 자동 병합기", layout="wide")

st.title("🏢 북경 치과전 비품 취합 자동화 (취합파일 양식 완전 대응 + 합계 + 가격 포함)")

uploaded_files = st.file_uploader("비품 주문서 파일 업로드 (여러 개 선택)", type=["xlsx"], accept_multiple_files=True)

# 품목별 단가 설정 (예시)
ITEM_PRICES = {
    '의자': 10000,
    '책상': 20000,
    '쇼케이스': 30000,
    '캐비닛': 15000,
    '인포데스크': 25000,
    '인포데스크/쇼케이스/캐비닛': 25000,  # 대표 품목 하나로 단가 설정
}

@st.cache_data
def process_file_full(file):
    try:
        df = pd.read_excel(file, sheet_name='1부스', header=None)

        # 메타 정보 추출
        company = df.iloc[7, 1] if not pd.isna(df.iloc[7, 1]) else "업체명 미기재"
        manager = df.iloc[7, 4] if not pd.isna(df.iloc[7, 4]) else ""
        booth_no = df.iloc[8, 4] if not pd.isna(df.iloc[8, 4]) else ""
        phone = df.iloc[9, 1] if not pd.isna(df.iloc[9, 1]) else ""
        email = df.iloc[8, 1] if not pd.isna(df.iloc[8, 1]) else ""
        memo = df.iloc[16, 5] if not pd.isna(df.iloc[16, 5]) else ""

        # 품목 테이블 (행 17~36, 열 0,2,4)
        temp_df = df.iloc[17:36, [0, 2, 4]].copy()
        temp_df.columns = ['품목', '기본제공수량', '최종기재수량']
        temp_df = temp_df.dropna(subset=['품목'])

        # 품목 처리
        expanded_rows = []
        for _, row in temp_df.iterrows():
            item = row['품목']
            qty = row['최종기재수량']
            if isinstance(qty, str) and any(k in qty for k in ['인포데스크', '쇼케이스', '캐비닛']):
                combined_qty = 0
                matches = re.findall(r'(인포데스크|쇼케이스|캐비닛)\s*\(\s*(\d+)\s*\)', qty)
                for _, count in matches:
                    combined_qty += int(count)
                expanded_rows.append({'ITEM': '인포데스크/쇼케이스/캐비닛', '수량': combined_qty})
            else:
                def extract_sum(x):
                    if isinstance(x, str):
                        nums = re.findall(r'\d+', x)
                        return sum(map(int, nums)) if nums else 0
                    if pd.isna(x):
                        return 0
                    return int(x)
                expanded_rows.append({'ITEM': item, '수량': extract_sum(qty)})

        expanded_df = pd.DataFrame(expanded_rows)
        expanded_df['가격'] = expanded_df['ITEM'].apply(lambda x: ITEM_PRICES.get(x, 0))
        expanded_df['합계'] = expanded_df['수량'] * expanded_df['가격']

        # 품목 수량 피벗
        item_dict = dict(zip(expanded_df['ITEM'], expanded_df['수량']))
        item_df = pd.DataFrame([item_dict])

        # 총 합계
        total_sum = expanded_df['합계'].sum()
        item_df['총합계'] = total_sum

        # 메타 정보
        meta = pd.DataFrame({
            'booth NO': [booth_no],
            'company name': [company],
            '담당자': [manager],
            '연락처': [phone],
            '이메일': [email],
            '비고': [memo],
        })

        full_row = pd.concat([meta, item_df], axis=1)
        return full_row, expanded_df

    except Exception as e:
        st.error(f"{file.name} 처리 중 오류 발생: {e}")
        return None, None

if uploaded_files:
    st.info("잠시만 기다려주세요. 업로드한 파일을 처리 중입니다...")

    result_rows = []
    detail_rows = []
    for file in uploaded_files:
        row, detail = process_file_full(file)
        if row is not None:
            result_rows.append(row)
        if detail is not None:
            detail['업체명'] = file.name.replace('.xlsx', '')  # 파일명으로 구분
            detail_rows.append(detail)

    if result_rows:
        final_result = pd.concat(result_rows, ignore_index=True)
        detail_result = pd.concat(detail_rows, ignore_index=True)

        st.success("✅ 모든 파일 처리 완료!")
        st.subheader("📋 전체 취합 결과")
        st.dataframe(final_result)

        st.subheader("📦 ITEM 상세 내역")
        st.dataframe(detail_result[['업체명', 'ITEM', '수량', '가격', '합계']])

        def to_excel(df1, df2):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df1.to_excel(writer, index=False, sheet_name='취합결과')
                df2.to_excel(writer, index=False, sheet_name='상세내역')
            return output.getvalue()

        st.download_button(
            label="📥 비품 취합 + 상세내역 다운로드",
            data=to_excel(final_result, detail_result),
            file_name="북경치과전_비품_최종취합_상세포함.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("파일에서 유효한 데이터를 찾을 수 없습니다.")
else:
    st.info("왼쪽에서 .xlsx 파일들을 업로드해 주세요. (1부스 시트 기준)")
