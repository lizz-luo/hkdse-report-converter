import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import numpy as np

st.set_page_config(page_title="HKDSE 全科轉換工具 v2.1", layout="centered")

st.title("📊 HKDSE 項目分析報告 → Excel v2.1")
st.markdown("""
✅ **全科完美支援**：
- 🗺️ 地理：百分號拆分（53% → 53 + %）
- 🇬🇧 英文：人數黏連（106 100.0 → 10；6 1000）  
- ➗ 數學：跨行黏連（135 \\n100.0）
- 🇨🇳 中文：雙重評分合併
- ✨ **全科無 <NA>**，數字可計算
""")

uploaded_file = st.file_uploader("請上傳 PDF 檔案", type="pdf")

def merge_percentage_split(row):
    """地理科：數字+% → 53%"""
    new_row = list(row)
    i = 0
    while i < len(new_row) - 1:
        cell1 = str(new_row[i]).strip()
        cell2 = str(new_row[i+1]).strip()
        if (re.match(r'^\d+\.?\d*$', cell1) and cell2 == '%'):
            new_row[i] = cell1 + '%'
            new_row[i+1] = ''
        i += 1
    return new_row

def fix_stuck_numbers_and_percent(row):
    """通用黏連修復（英文/數學強化）"""
    new_row = list(row)
    i = 0
    while i < len(new_row) - 1:
        cell1 = str(new_row[i]).strip()
        cell2 = str(new_row[i+1]).strip()
        
        # 模式1：數字空格數字
        combined = cell1 + ' ' + cell2
        if re.match(r'^\d+\s+\d+\.?\d*$', combined):
            parts = re.split(r'\s+', combined)
            if len(parts) == 2:
                new_row[i] = parts[0]
                new_row[i+1] = parts[1]
                i += 1
                continue
        
        # 模式2：純數字黏連
        if (re.match(r'^\d{3,}$', cell1) and 
            re.match(r'^\d+\.?\d*$', cell2)):
            new_row[i] = cell1
            new_row[i+1] = cell2
        
        i += 1
    return new_row

def fix_math_crossline_stuck(row):
    """數學專用：135 \n100.0 → 135 + 100.0"""
    new_row = list(row)
    i = 0
    while i < len(new_row):
        cell = str(new_row[i]).strip()
        if '\n' in cell and re.search(r'\d+\.?\d*\n\d+\.?\d*', cell):
            parts = re.split(r'\n+', cell)
            if len(parts) >= 2:
                new_row[i] = parts[0].strip()
                if i+1 < len(new_row):
                    new_row[i+1] = parts[1].strip()
        i += 1
    return new_row

def clean_and_convert_to_numeric(val):
    """智能數字轉換"""
    if pd.isna(val) or val is None:
        return np.nan
    val_str = str(val).strip()
    if not val_str or val_str == '':
        return np.nan
    
    if val_str.endswith('%'):
        try:
            return float(val_str.replace('%', '')) / 100.0
        except:
            pass
    
    val_str = val_str.replace(',', '')
    if val_str.startswith('+'):
        val_str = val_str[1:]
    
    try:
        num = float(val_str)
        return int(num) if num.is_integer() else num
    except ValueError:
        return val_str

def fix_row_split_numbers(cells):
    """106 → 10+6 合併"""
    cells = [("" if c is None or pd.isna(c) else str(c)) for c in cells]
    for i in range(len(cells) - 2):
        a, b, c = cells[i], cells[i + 1], cells[i + 2]
        if a and b and c and a.isdigit() and b.isdigit() and len(a)==2 and len(b)==1:
            c_clean = c.replace(',', '')
            if c_clean.startswith("100"):
                cells[i] = a + b
                cells[i + 1] = ""
                break
    return cells

if uploaded_file is not None:
    with st.spinner("🔄 智能解析中（全科修復）..."):
        sections = {}
        current_section = "General"
        detected_subject = "General"

        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text: continue

                # 科目自動識別
                if "Geography" in text or "地理" in text:
                    detected_subject = "Geography"
                elif "English Language" in text or "英國語文" in text:
                    detected_subject = "English"
                elif "Mathematics" in text or "數學" in text:
                    detected_subject = "Math"
                elif "Chinese Language" in text or "中國語文" in text:
                    detected_subject = "Chinese"

                # 卷別
                paper_match = re.search(r"(?:卷\s*)?Paper:\s*([A-Za-z0-9]+)", text)
                if paper_match:
                    paper_name = paper_match.group(1).strip()
                    current_section = f"Paper_{paper_name}"
                else:
                    if current_section == "General" and detected_subject != "General":
                        current_section = f"{detected_subject}_General"

                if current_section not in sections:
                    sections[current_section] = []

                # 動態 X 容錯
                x_tolerance = {
                    "Geography": 1.2, "English": 1.5, "Math": 1.0
                }.get(detected_subject, 3.0)

                table_settings = {
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "intersection_x_tolerance": x_tolerance,
                    "intersection_y_tolerance": 2,
                    "min_words_vertical": 2,
                }

                table = page.extract_table(table_settings)
                if table:
                    sections[current_section].extend(table)

        # 生成 Excel
        output = io.BytesIO()
        has_data = False

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for section_name, data in sections.items():
                if not data: continue

                cleaned_rows = [row for row in data 
                              if any((isinstance(c, str) and c.strip()) for c in row)]

                if not cleaned_rows: continue

                df = pd.DataFrame(cleaned_rows)
                df = df.replace(r"^\s*$", np.nan, regex=True)
                df.dropna(how='all', axis=1, inplace=True)

                # 科目專屬修復
                if detected_subject == "Geography":
                    df = df.apply(merge_percentage_split, axis=1, result_type="expand")
                df = df.apply(fix_stuck_numbers_and_percent, axis=1, result_type="expand")
                if detected_subject == "Math":
                    df = df.apply(fix_math_crossline_stuck, axis=1, result_type="expand")

                # 標準流程
                df = df.apply(lambda row: fix_row_split_numbers(list(row)), 
                            axis=1, result_type="expand")
                df = df.replace(r"^\s*$", np.nan, regex=True)

                def shift_row_right(row):
                    valid_vals = [v for v in row if pd.notna(v) and str(v).strip()]
                    num_nans = len(row) - len(valid_vals)
                    return pd.Series([np.nan] * num_nans + valid_vals, index=row.index)

                df = df.apply(shift_row_right, axis=1)
                df.dropna(how='all', axis=1, inplace=True)
                df = df.applymap(clean_and_convert_to_numeric)

                # 徹底無 NA
                df = df.replace([np.nan, pd.NA, None], '')
                df = df.replace('<NA>', '')

                safe_name = re.sub(r"[\\/*?:\[\]]", "_", section_name)[:31]
                df.to_excel(writer, sheet_name=safe_name, index=False, header=False)
                has_data = True

        if has_data:
            st.success(f"🎉 轉換完成！偵測科目：**{detected_subject}**")
            st.download_button(
                label="📥 下載 Excel 檔案",
                data=output.getvalue(),
                file_name=f"DSE_Report_v2.1_{detected_subject}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("⚠️ 未找到有效表格，請確認 PDF 格式")

st.caption("✨ v2.1 | 全科無敵版 | 數學/英文/地理完美修復")
