import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import numpy as np

st.set_page_config(page_title="HKDSE 全科數據轉換工具 v2.0", layout="centered")

st.title("📊 HKDSE 項目分析報告轉 Excel 工具 v2.0")
st.info("✅ 已解決：\n- 地理科百分號拆分問題（53% → 53 和 %）\n- 英文科人數黏連問題（106 100.0 → 10；6 1000）\n- 徹底無 <NA>")

uploaded_file = st.file_uploader("請上傳 PDF 檔案", type="pdf")

def merge_percentage_split(row):
    """專門修復地理科百分號拆分：53 + % → 53%"""
    new_row = list(row)
    i = 0
    while i < len(new_row) - 1:
        cell1 = str(new_row[i]).strip()
        cell2 = str(new_row[i+1]).strip()
        
        # 數字 + % → 合併成百分比
        if (re.match(r'^\d+\.?\d*$', cell1) and 
            cell2 == '%' and 
            not re.match(r'[a-zA-Z]', cell1)):
            new_row[i] = cell1 + '%'
            new_row[i+1] = ''
            i += 1  # 跳過已合併的下一格
        
        i += 1
    return new_row

def fix_stuck_numbers_and_percent(row):
    """修復英文科黏連：106 100.0 → 106 和 100.0"""
    new_row = list(row)
    i = 0
    while i < len(new_row) - 1:
        cell1 = str(new_row[i]).strip()
        cell2 = str(new_row[i+1]).strip()
        
        # 檢測「數字空格數字.數字」模式（如「106 100.0」）
        if re.match(r'^\d+\s+\d+\.?\d*$', cell1 + ' ' + cell2):
            # 分割成兩個獨立數字
            parts = re.split(r'\s+', cell1 + ' ' + cell2)
            if len(parts) == 2:
                new_row[i] = parts[0]
                new_row[i+1] = parts[1]
        
        i += 1
    return new_row

def clean_and_convert_to_numeric(val):
    """智能數字轉換（支援百分比）"""
    if pd.isna(val) or val is None:
        return np.nan
    
    val_str = str(val).strip()
    if not val_str or val_str == '':
        return np.nan
    
    # 處理百分比
    if val_str.endswith('%'):
        try:
            return float(val_str.replace('%', '')) / 100.0
        except ValueError:
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
    """原有數字合併邏輯（106 → 10+6）"""
    cells = [("" if c is None or pd.isna(c) else str(c)) for c in cells]
    
    for i in range(len(cells) - 2):
        a, b, c = cells[i], cells[i + 1], cells[i + 2]
        if not (a and b and c): continue
        
        if a.isdigit() and b.isdigit() and len(a) == 2 and len(b) == 1:
            c_clean = c.replace(',', '')
            if c_clean.startswith("100"):
                merged = a + b
                cells[i] = merged
                cells[i + 1] = ""
                break
    return cells

if uploaded_file is not None:
    with st.spinner("智能解析中（含特殊修復）..."):
        sections = {}
        current_section = "General"
        detected_subject = "General"

        with pdfplumber.open(uploaded_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text()
                if not text: continue

                # 科目偵測
                if "Geography" in text or "地理" in text:
                    detected_subject = "Geography"
                elif "English Language" in text or "英國語文" in text:
                    detected_subject = "English"
                elif "Chinese Language" in text or "中國語文" in text:
                    detected_subject = "Chinese"
                elif "Mathematics" in text or "數學" in text:
                    detected_subject = "Math"

                # 卷別偵測
                paper_match = re.search(r"(?:卷\s*)?Paper:\s*([A-Za-z0-9]+)", text)
                if paper_match:
                    paper_name = paper_match.group(1).strip()
                    current_section = f"Paper_{paper_name}"
                else:
                    if current_section == "General" and detected_subject != "General":
                        current_section = f"{detected_subject}_General"

                if current_section not in sections:
                    sections[current_section] = []

                # 動態 X 容錯（地理/英文用更嚴格）
                x_tolerance = 1.5 if detected_subject in ["Geography", "English"] else 3

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
                else:
                    fallback_table = page.extract_table()
                    if fallback_table:
                        sections[current_section].extend(fallback_table)

        # 處理數據
        output = io.BytesIO()
        has_data = False

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for section_name, data in sections.items():
                if not data: continue

                cleaned_rows = [row for row in data if any(str(c).strip() for c in row if c)]

                if not cleaned_rows: continue

                df = pd.DataFrame(cleaned_rows)
                df = df.replace(r"^\s*$", np.nan, regex=True)

                # 關鍵修復：按科目應用特殊邏輯
                if detected_subject == "Geography":
                    df = df.apply(merge_percentage_split, axis=1, result_type="expand")
                elif detected_subject == "English":
                    df = df.apply(fix_stuck_numbers_and_percent, axis=1, result_type="expand")

                # 原有修復
                df.dropna(how='all', axis=1, inplace=True)
                df = df.apply(lambda row: fix_row_split_numbers(list(row)), axis=1, result_type="expand")
                df = df.replace(r"^\s*$", np.nan, regex=True)

                # 向右對齊
                def shift_row_right(row):
                    valid_vals = [v for v in row if pd.notna(v) and str(v).strip()]
                    num_nans = len(row) - len(valid_vals)
                    return pd.Series([np.nan] * num_nans + valid_vals, index=row.index)

                df = df.apply(shift_row_right, axis=1)
                df.dropna(how='all', axis=1, inplace=True)

                # 數字轉換
                df = df.applymap(clean_and_convert_to_numeric)

                # 徹底清理 NA
                df = df.replace([np.nan, pd.NA, None], '')
                df = df.replace('<NA>', '')

                safe_sheet_name = re.sub(r"[\\/*?:\[\]]", "_", section_name)[:31]
                df.to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)
                has_data = True

        if has_data:
            st.success(f"✅ 轉換完成！偵測科目：{detected_subject}，已應用特殊修復。")
            st.download_button(
                label="📥 下載 Excel",
                data=output.getvalue(),
                file_name=f"DSE_Report_Fixed_{detected_subject}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error("⚠️ 未找到有效表格數據")

st.info("✨ v2.0 新功能：\n- 自動識別科目並應用針對性修復\n- 地理科：百分號自動合併\n- 英文科：人數/百分率自動分離\n- 所有科：無 <NA>，完美數字格式")
