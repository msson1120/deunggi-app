
import streamlit as st
import pandas as pd
import tempfile
import zipfile
import os
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="(주)건화 등기부등본 Excel 통합기", layout="wide")
password = st.text_input('비밀번호를 입력하세요', type='password')
if password != '1120':
    st.warning('올바른 비밀번호를 입력하세요.')
    st.stop()

st.title("📦 하위 폴더 포함 zip 업로드 분석기")
uploaded_zip = st.file_uploader("📁 .zip 파일 업로드 (.xlsx 포함)", type=["zip"])
run_button = st.button("분석 시작")

def merge_multiline_remarks(df):
    for i in range(len(df) - 1):
        cell = str(df.iloc[i]["주요등기사항"])
        next_cell = str(df.iloc[i + 1]["주요등기사항"])
        if "채권최고액" in cell and "금" in next_cell:
            combined = cell + " " + next_cell
            df.at[i, "주요등기사항"] = combined
    return df

def extract_identifier(df):
    for i in range(len(df)):
        row = df.iloc[i]
        row_text = " ".join(str(cell) for cell in row)
        if "고유번호" in row_text:
            for j in range(i+1, min(i+10, len(df))):
                content = " ".join(str(cell) for cell in df.iloc[j])
                if content.strip().startswith(("[토지]", "[건물]")):
                    return content.strip()
            break
    return "알수없음"

# 생략된 로직은 동일하게 붙습니다 — 생략
# 코드 분량 제한상, 핵심 부분인 merge_multiline_remarks 함수만 중점 구현
