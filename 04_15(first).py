import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="비품 수량 자동 병합기", layout="wide")
st.title("🏢 북경 치과전 비품 취합 자동화 (1+2항목 완전 대응)")

uploaded_files = st.file_uploader(
    "비품 주문서 파일 업로드 (여러 개 선택)", type=["xlsx"], accept_multiple_files=True
)

@st.cache_data
def process_file_full(file):
    try:
        df = pd.read_excel(file, sheet_name='1부스', header=None)

        company = df.iloc[7, 1] if not pd.isna(df.iloc[7, 1]) else "업체명 미기재"
        manager = df.iloc[7, 4] if not pd.isna(df.iloc[7, 4]) else ""
        booth_no = df.iloc[8, 4] if not pd.isna(df.iloc[8, 4]) else ""
        phone = df.iloc[9, 1] if not pd.isna(df.iloc[9, 1]) else ""
        email = df.iloc[8, 1] if not pd.isna(df.iloc[8, 1]) else ""
        memo = df.iloc[16, 5] if not pd.isna(df.iloc[16, 5]) else ""

        # 기본 비품: 17~36행
        temp_df = df.iloc[17:36, [0, 2, 4]].copy()
        temp_df.columns = ['품목', '기본제공수량', '최종기재수량']
        temp_df = temp_df.dropna(subset=['품목'])

        expanded_rows = []
        for _, row in temp_df.iterrows():
            item = row['품목']
            qty = row['최종기재수량']
            if isinstance(qty, str) and any(k in qty for k in ['인포데스크', '쇼케이스', '캐비닛']):
                matches = re.findall(r'(인포데스크|쇼케이스|캐비닛)\s*\(\s*(\d+)\s*\)', qty)
                for item_name, count in matches:
                    expanded_rows.append({'ITEM': f"{item_name}", '수량': int(count)})
            else:
                def extract_sum(x):
                    if isinstance(x, str):
                        nums = re.findall(r'\d+', x)
                        return sum(map(int, nums)) if nums else 0
                    if pd.isna(x): return 0
                    return int(x)
                expanded_rows.append({'ITEM': item, '수량': extract_sum(qty)})

        basic_df = pd.DataFrame(expanded_rows)
        basic_df['가격'] = 0
        basic_df['합계'] = 0
        basic_df['비고'] = ""

        # 추가 비품 (A33 기준 = index 32)
        additional_rows = []
        for i in range(32, df.shape[0]):
            item = df.iloc[i, 0]
            qty = df.iloc[i, 2]
            price = df.iloc[i, 3]
            memo_ = df.iloc[i, 5]
            if pd.isna(item) or str(item).strip() == "":
                continue
            try:
                qty = int(qty) if not pd.isna(qty) else 0
                price = int(str(price).replace(",", "")) if not pd.isna(price) else 0
                additional_rows.append({
                    'ITEM': str(item).strip(),
                    '수량': qty,
                    '가격': price,
                    '합계': qty * price,
                    '비고': memo_ if not pd.isna(memo_) else ""
                })
            except:
                continue

        additional_df = pd.DataFrame(additional_rows)

        # 병합 및 그룹화
        all_items = pd.concat([basic_df, additional_df], ignore_index=True)
        grouped = all_items.groupby("ITEM", as_index=False).agg({
            "수량": "sum",
            "가격": "sum",
            "합계": "sum",
            "비고": lambda x: " / ".join(set(x.dropna().astype(str))) if not x.isna().all() else ""
        })

        grouped.insert(0, "업체명", company)
        return grouped

    except Exception as e:
        st.error(f"{file.name} 처리 중 오류 발생: {e}")
        return None

if uploaded_files:
    st.info("잠시만 기다려주세요. 업로드한 파일을 처리 중입니다...")
    result_rows = []
    for file in uploaded_files:
        row = process_file_full(file)
        if row is not None:
            result_rows.append(row)

    if result_rows:
        final_result = pd.concat(result_rows, ignore_index=True)

        st.success("✅ 모든 파일 처리 완료!")
        st.dataframe(final_result)

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='취합결과', startrow=0, startcol=0)
            return output.getvalue()

        st.download_button(
            label="📥 최종 취합 파일 다운로드",
            data=to_excel(final_result),
            file_name="북경치과전_비품_취합_완성본.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠ 처리 가능한 데이터가 없습니다.")
else:
    st.info("왼쪽에서 .xlsx 파일을 하나 이상 업로드해 주세요.")
