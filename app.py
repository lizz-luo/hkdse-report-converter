import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import numpy as np

st.set_page_config(page_title="HKDSE 全科數據轉換工具", layout="centered")

st.title("📊 HKDSE 項目分析報告轉 Excel 工具")
st.write("請上傳考評局的 PDF 報告，系統會自動提取、向右對齊、刪除多餘空欄，並**將成績轉為真實數字格式**以便計算。")

uploaded_file = st.file_uploader("請點擊或拖曳上傳 PDF 檔案", type="pdf")

custom_table_settings = {
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
    "intersection_x_tolerance": 2,
    "intersection_y_tolerance": 2,
    "min_words_vertical": 2
}

def clean_and_convert_to_numeric(val):
    """將文字轉換為真實數字的智能函數"""
    if pd.isna(val):
        return val
        
    val_str = str(val).strip()
    
    # 移除千位數逗號
    val_str = val_str.replace(',', '')
    
    # 如果是百分比 (例如 "85%")，去掉 % 並除以 100 變成 0.85
    if val_str.endswith('%'):
        try:
            return float(val_str.replace('%', '')) / 100.0
        except ValueError:
            pass
            
    # 如果帶有加號 (例如 "+0.25")，去掉加號轉換為正數
    if val_str.startswith('+'):
        val_str = val_str[1:]
        
    # 嘗試轉換為浮點數或整數
    try:
        # 如果小數點後為 0，就轉為整數；否則轉浮點數
        num = float(val_str)
        return int(num) if num.is_integer() else num
    except ValueError:
        return val_str # 如果轉換失敗（例如是純文字的題號），就保留原樣

if uploaded_file is not None:
    with st.spinner("正在智能解析 PDF 並轉換數字格式，這可能需要幾秒鐘..."):
        sections = {}
        current_section = "General"

        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                
                paper_match = re.search(r'(?:卷\s*)?Paper:\s*([A-Za-z0-9]+)', text)
                if paper_match:
                    paper_name = paper_match.group(1).strip()
                    current_section = f"Paper_{paper_name}"
                    if current_section not in sections:
                        sections[current_section] = []
                elif "Mathematics" in text and current_section == "General":
                    current_section = "Math_Compulsory"
                elif "English" in text and current_section == "General":
                    current_section = "Eng_Lang"
                elif "Chinese" in text and current_section == "General":
                    current_section = "Chi_Lang"
                
                if current_section not in sections:
                    sections[current_section] = []
                    
                table = page.extract_table(custom_table_settings)
                
                if table:
                    sections[current_section].extend(table)
                else:
                    fallback_table = page.extract_table()
                    if fallback_table:
                        sections[current_section].extend(fallback_table)

        output = io.BytesIO()
        has_data = False
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for section_name, data in sections.items():
                if data:
                    clean_data = [row for row in data if any(cell.strip() for cell in row if isinstance(cell, str))]
                    if clean_data:
                        df = pd.DataFrame(clean_data)
                        
                        # 1. 統天空值處理
                        df.replace(r'^\s*$', pd.NA, regex=True, inplace=True)
                        df.dropna(how='all', axis=1, inplace=True)
                        
                        # 2. 向右對齊邏輯
                        def shift_row_right(row):
                            valid_values = row.dropna().tolist()
                            num_nans_to_add = len(row) - len(valid_values)
                            shifted_values = [pd.NA] * num_nans_to_add + valid_values
                            return pd.Series(shifted_values, index=row.index)
                        
                        df = df.apply(shift_row_right, axis=1)
                        
                        # 3. 再次刪除因向右對齊而產生的左側「全空欄」
                        df.dropna(how='all', axis=1, inplace=True)
                        
                        # 4. 套用數字轉換函數到整個 DataFrame 的每一個儲存格
                        df = df.applymap(clean_and_convert_to_numeric)
                        
                        safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', section_name)[:31]
                        df.to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)
                        has_data = True

        if has_data:
            st.success("✅ 智能轉換成功！數據已對齊並轉為數字格式。請點擊下方按鈕下載 Excel。")
            st.download_button(
                label="📥 下載最終版 Excel 檔案",
                data=output.getvalue(),
                file_name="DSE_Report_Numeric.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("⚠️ 未能從上傳的 PDF 中找到有效的成績表格數據。")
