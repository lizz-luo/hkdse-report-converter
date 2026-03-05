import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import numpy as np

st.set_page_config(page_title="HKDSE 全科數據轉換工具", layout="centered")

st.title("📊 HKDSE 項目分析報告轉 Excel 工具")
st.write("請上傳考評局的 PDF 報告（支援中英數等各科），系統會自動提取數據、清理空白欄位，並加上標準表頭。")

uploaded_file = st.file_uploader("請點擊或拖曳上傳 PDF 檔案", type="pdf")

# 進階 PDF 提取參數（針對無網格線的表格）
custom_table_settings = {
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
    "intersection_x_tolerance": 2,
    "intersection_y_tolerance": 2,
    "min_words_vertical": 2
}

# 您指定的標準表頭
standard_headers = [
    "Item", 
    "Max Mark", 
    "Your school Attm. No.", 
    "Your school Attem. %", 
    "Your school Mean", 
    "Your school Mean %", 
    "Your school SD", 
    "Day schools Attem. %", 
    "Day schools Mean", 
    "Day schools Mean %", 
    "Day schools SD"
]

if uploaded_file is not None:
    with st.spinner("正在智能解析並清洗數據，這可能需要幾秒鐘..."):
        sections = {}
        current_section = "General"

        # 第一階段：讀取 PDF 數據
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                
                # 智能捕捉卷別
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

        # 第二階段：數據清洗與寫入 Excel
        output = io.BytesIO()
        has_data = False
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for section_name, data in sections.items():
                if data:
                    # 將二維陣列轉為 Pandas DataFrame
                    df = pd.DataFrame(data)
                    
                    # 1. 將所有的 None、空字串、或只有空格的字串轉為 NaN
                    df.replace(r'^\s*$', np.nan, regex=True, inplace=True)
                    
                    # 2. 刪除所有「完全空白（全為 NaN）」的橫列 (rows) 和直欄 (columns)
                    df.dropna(how='all', axis=0, inplace=True)
                    df.dropna(how='all', axis=1, inplace=True)
                    
                    # 3. 找出真正的數據行（排除考評局原有的亂碼表頭）
                    # 邏輯：考評局數據行通常包含數字、小數點或加減號，而表頭都是純文字
                    # 我們過濾掉那些「欄位含有中文字或特定英文字」的標題行
                    def is_data_row(row):
                        row_str = ' '.join([str(x) for x in row if pd.notna(x)])
                        return bool(re.search(r'\d', row_str)) and not "Mean" in row_str and not "平均" in row_str
                    
                    df = df[df.apply(is_data_row, axis=1)]
                    
                    if not df.empty:
                        # 4. 重新對齊欄位與套用標準表頭
                        # 考評局除了第一欄(題號)外，通常有 10 個數據欄位，加上差值欄位可能會有 12-14 欄
                        # 如果清洗後的欄位數剛好大於等於我們的標準表頭數(11)，我們就抓取前 11 欄並套用表頭
                        if len(df.columns) >= len(standard_headers):
                            # 只保留前 11 欄核心數據（捨棄最後面的「Diff差距」等用不到的欄位）
                            df = df.iloc[:, :len(standard_headers)]
                            df.columns = standard_headers
                            # 將設定好表頭的數據寫入 Excel
                            safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', section_name)[:31]
                            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                            has_data = True
                        else:
                            # 如果因為 PDF 排版異常導致欄位數不足，就保持無表頭匯出，以免報錯
                            safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', section_name)[:31]
                            df.to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)
                            has_data = True

        if has_data:
            st.success("✅ 數據清洗與表頭套用成功！請點擊下方按鈕下載 Excel。")
            st.download_button(
                label="📥 下載乾淨的 Excel 檔案",
                data=output.getvalue(),
                file_name="DSE_Report_Cleaned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("⚠️ 未能提取有效數據。")
