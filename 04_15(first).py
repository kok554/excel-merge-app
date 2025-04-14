import streamlit as st
import pandas as pd
from io import BytesIO

st.title("엑셀 통합 앱")

uploaded_a = st.file_uploader("A 파일 업로드", type=["xlsx"], key="a")
uploaded_b = st.file_uploader("B 파일 업로드", type=["xlsx"], key="b")
uploaded_c = st.file_uploader("C 파일 업로드", type=["xlsx"], key="c")

if uploaded_a and uploaded_b and uploaded_c:
    df_a = pd.read_excel(uploaded_a)
    df_b = pd.read_excel(uploaded_b)
    df_c = pd.read_excel(uploaded_c)

    merged_df = pd.concat([df_a, df_b, df_c], ignore_index=True)

    st.subheader("병합된 결과")
    st.dataframe(merged_df)

    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Merged')
        processed_data = output.getvalue()
        return processed_data

    excel_data = to_excel(merged_df)

    st.download_button(
        label="통합 엑셀 다운로드",
        data=excel_data,
        file_name="D_merged.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
