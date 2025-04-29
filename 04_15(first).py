import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="북경치과전 비품 자동 취합기", layout="wide")
st.title("📦 북경치과전 비품 주문서 자동 병합기")

uploaded_files = st.file_uploader("🧾 참가업체 엑셀(.xlsx) 업로드", type=["xlsx"], accept_multiple_files=True)

@st.cache_data
def extract_info_from_file(file):
    try:
        df = pd.read_excel(file, sheet_name="1부스", header=None)

        company = df.iloc[7, 1] if not pd.isna(df.iloc[7, 1]) else "업체명 미기재"
        manager = df.iloc[7, 4] if not pd.isna(df.iloc[7, 4]) else ""
        booth_no = df.iloc[8, 4] if not pd.isna(df.iloc[8, 4]) else ""
        phone = df.iloc[9, 1] if not pd.isna(df.iloc[9, 1]) else ""
        email = df.iloc[8, 1] if not pd.isna(df.iloc[8, 1]) else ""
        memo = df.iloc[16, 5] if not pd.isna(df.iloc[16, 5]) else ""

        # 기본 비품: 17~36행
        default_df = df.iloc[17:36, [0, 4]].copy()
        default_df.columns = ['ITEM', 'QTY']
        default_df = default_df.dropna(subset=['ITEM'])

        # 추가 비품: A33 기준 (index 32부터)
        additional = []
        for i in range(32, df.shape[0]):
            item = df.iloc[i, 0]
            qty = df.iloc[i, 2]
            if pd.isna(item) or str(item).strip() == "":
                continue
            try:
                item = str(item).strip()
                qty = int(qty) if not pd.isna(qty) else 0
                additional.append({'ITEM': item, 'QTY': qty})
            except:
                continue

        combined_df = pd.concat([default_df, pd.DataFrame(additional)], ignore_index=True)

        # 특수 항목 분리
        expanded_rows = []
        for _, row in combined_df.iterrows():
            item = row['ITEM']
            qty = row['QTY']
            if isinstance(qty, str) and any(k in qty for k in ['인포데스크', '쇼케이스', '캐비닛']):
                matches = re.findall(r'(인포데스크|쇼케이스|캐비닛)\s*\(?\s*(\d+)\s*\)?', qty)
                for sub_item, count in matches:
                    expanded_rows.append((sub_item, int(count)))
            elif isinstance(item, str) and re.search(r'\(.+\)', item):
                matches = re.findall(r'([가-힣A-Za-z]+)\s*\(?\s*(\d+)\s*\)?', item)
                for sub_item, count in matches:
                    expanded_rows.append((sub_item, int(count)))
            else:
                try:
                    expanded_rows.append((str(item).strip(), int(qty)))
                except:
                    continue

        # 집계
        item_df = pd.DataFrame(expanded_rows, columns=["ITEM", "수량"])
        item_summary = item_df.groupby("ITEM").sum(numeric_only=True).T
        item_summary["합계"] = item_summary.sum(axis=1)

        # 메타정보 붙이기
        item_summary.insert(0, "회사명", company)
        item_summary.insert(1, "부스번호", booth_no)
        item_summary["연락처"] = phone
        item_summary["이메일"] = email
        item_summary["담당자"] = manager
        item_summary["비고"] = memo

        return item_summary

    except Exception as e:
        st.error(f"{file.name} 처리 중 오류: {e}")
        return None

if uploaded_files:
    all_results = []
    for file in uploaded_files:
        res = extract_info_from_file(file)
        if res is not None:
            all_results.append(res)

    if all_results:
        merged = pd.concat(all_results, ignore_index=True)
        st.success("✅ 모든 파일 병합 완료")
        st.dataframe(merged)

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="북경치과전_비품_취합")
            return output.getvalue()

        st.download_button(
            label="📥 비품 취합 엑셀 다운로드",
            data=to_excel(merged),
            file_name="북경치과전_비품_취합_최종.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠ 처리 가능한 데이터가 없습니다.")
else:
    st.info("📤 좌측에서 참가업체 엑셀 파일을 업로드하세요.")
