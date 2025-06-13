
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

st.title("📦 폴더 포함 zip 파일 업로드 분석기")
st.markdown("""
압축파일(.zip) 안의 폴더 구조와 관계없이 모든 엑셀 파일을 자동 분석합니다.
""")

uploaded_zip = st.file_uploader("📁 .zip 파일을 업로드하세요 (내부에 .xlsx 파일 포함)", type=["zip"])
run_button = st.button("분석 시작")

def trim_after_reference_note(df):
    for i, row in df.iterrows():
        row_text = "".join(str(cell) for cell in row)
        normalized = re.sub(r"\s+", "", row_text)
        if "참고사항" in normalized or "참고" in normalized or "비고" in normalized:
            return df.iloc[:i]
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

def keyword_match_partial(cell, keyword):
    if pd.isnull(cell): return False
    return keyword.replace(" ", "") in str(cell).replace(" ", "")

def keyword_match_exact(cell, keyword):
    if pd.isnull(cell): return False
    return re.sub(r"\s+", "", str(cell)) == re.sub(r"\s+", "", keyword)

def extract_section_range(df, start_kw, end_kw_list, match_fn):
    df = df.fillna("")
    df.columns = range(df.shape[1])
    start_idx, end_idx = None, len(df)
    for i, row in df.iterrows():
        if any(match_fn(cell, start_kw) for cell in row):
            start_idx = i + 1
            break
    if start_idx is None:
        return pd.DataFrame(), False
    for i in range(start_idx, len(df)):
        row = df.iloc[i]
        if any(any(match_fn(cell, end_kw) for cell in row) for end_kw in end_kw_list):
            end_idx = i
            break
    section = df.iloc[start_idx:end_idx].copy()
    is_empty = section.replace("", pd.NA).dropna(how="all").empty
    return section if not is_empty else pd.DataFrame([["기록없음"]]), not is_empty

def extract_named_cols(section, col_keywords):
    header_row = section.iloc[0]
    col_map = {}
    for target in col_keywords:
        for idx, val in header_row.items():
            if keyword_match_partial(val, target):
                col_map[target] = idx
                break
    rows = []
    for i in range(1, len(section)):
        row = section.iloc[i]
        rows.append({key: row.get(col_map.get(key), "") for key in col_keywords})
    return pd.DataFrame(rows)

def find_keyword_header(section, col_keywords, max_search_rows=15):
    section = section.fillna("").astype(str)
    for i in range(min(max_search_rows, len(section))):
        row = section.iloc[i]
        match_count = sum(any(keyword_match_exact(cell, kw) for cell in row) for kw in col_keywords)
        if match_count >= 3:
            return i, row
    return None, None

def find_col_index(header_row, keyword):
    for idx, val in header_row.items():
        if keyword_match_exact(val, keyword):
            return idx
    return None

def extract_precise_named_cols(section, col_keywords):
    header_idx, header_row = find_keyword_header(section, col_keywords)
    if header_idx is None:
        header_row = section.iloc[0]
        start_row = 1
    else:
        start_row = header_idx + 1
    col_map = {key: find_col_index(header_row, key) for key in col_keywords if find_col_index(header_row, key) is not None}
    if not col_map:
        return pd.DataFrame([["기록없음"]])
    rows = []
    for i in range(start_row, len(section)):
        row = section.iloc[i]
        row_dict = {key: row[col_map[key]] if col_map[key] in row else "" for key in col_map}
        rows.append(row_dict)
    return pd.DataFrame(rows)

if run_button and uploaded_zip:
    temp_dir = tempfile.mkdtemp()
    szj_list, syg_list, djg_list = [], [], []

    with zipfile.ZipFile(uploaded_zip, "r") as z:
        z.extractall(temp_dir)

    # ✅ 하위 폴더 포함 모든 .xlsx 탐색
    excel_files = []
    for root, _, files in os.walk(temp_dir):
        for f in files:
            if f.lower().endswith(".xlsx"):
                excel_files.append(os.path.join(root, f))

    for path in excel_files:
        try:
            xls = pd.ExcelFile(path)
            df = xls.parse(xls.sheet_names[0]).fillna("")
            name = extract_identifier(df)

            szj_sec, has_szj = extract_section_range(df, "소유지분현황", ["소유권", "저당권"], match_fn=keyword_match_partial)
            syg_sec, has_syg = extract_section_range(df, "소유권.*사항", ["저당권"], match_fn=keyword_match_exact)
            djg_sec, has_djg = extract_section_range(df, "3.(근)저당권및전세권등(을구)", ["참고", "비고", "총계", "전산자료"], match_fn=keyword_match_exact)

            if has_szj:
                szj_df = extract_named_cols(szj_sec, ["등기명의인", "(주민)등록번호", "최종지분", "주소", "순위번호"])
                szj_df.insert(0, "파일명", name)
                szj_list.append(szj_df)
            else:
                st.warning(f"{name} 파일에서 소유지분현황 데이터 없음")
                szj_list.append(pd.DataFrame([[name, "기록없음"]], columns=["파일명", "등기명의인"]))

            if has_syg:
                syg_df = extract_precise_named_cols(syg_sec, ["순위번호", "등기목적", "접수정보", "주요등기사항", "대상소유자"])
                syg_df.insert(0, "파일명", name)
                syg_list.append(syg_df)
            else:
                st.warning(f"{name} 파일에서 소유권사항 데이터 없음")
                syg_list.append(pd.DataFrame([[name, "기록없음"]], columns=["파일명", "순위번호"]))

            if has_djg:
                djg_df = extract_precise_named_cols(djg_sec, ["순위번호", "등기목적", "접수정보", "주요등기사항", "대상소유자"])
                djg_df = trim_after_reference_note(djg_df)
                djg_df.insert(0, "파일명", name)
                djg_list.append(djg_df)
            else:
                st.warning(f"{name} 파일에서 저당권사항 데이터 없음")
                djg_list.append(pd.DataFrame([[name, "기록없음"]], columns=["파일명", "순위번호"]))

        except Exception as e:
            st.warning(f"{os.path.basename(path)} 처리 중 오류 발생: {e}")

    wb = Workbook()
    for sheetname, data in zip(
        ["1. 소유지분현황 (갑구)", "2. 소유권사항 (갑구)", "3. 저당권사항 (을구)"],
        [szj_list, syg_list, djg_list]
    ):
        ws = wb.create_sheet(title=sheetname)
        if data:
            df = pd.concat(data, ignore_index=True)
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)
        else:
            ws.append(["기록없음"])

    wb.remove(wb["Sheet"])
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name)
        st.success("✅ 분석 완료! 다운로드 버튼을 클릭하세요.")
        st.download_button("📥 결과 다운로드", data=open(tmp.name, "rb"), file_name="등기사항_통합_시트별구성.xlsx")
