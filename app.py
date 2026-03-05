import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="HKDSE 全科數據轉換工具", layout="centered")

st.title("📊 HKDSE 項目分析報告轉 Excel 工具")
st.write("請上傳考評局的 PDF 報告（支援**所有科目**），系統會自動提取數據、消除空欄，並**將所有數據向右對齊**。")

uploaded_file = st.file_uploader("請點擊或拖曳上傳 PDF 檔案", type="pdf")

# 專門對付考評局無格線表格的進階參數
custom_table_settings = {
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
    "intersection_x_tolerance": 2,
    "intersection_y_tolerance": 2,
    "min_words_vertical": 2
}

if uploaded_file is not None:
    with st.spinner("正在智能解析 PDF 並向右對齊數據，這可能需要幾秒鐘..."):
        sections = {}
        current_section = "General"

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
                    
                # 提取表格
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
                    # 清洗空白行
                    clean_data = [row for row in data if any(cell.strip() for cell in row if isinstance(cell, str))]
                    if clean_data:
                        df = pd.DataFrame(clean_data)
                        
                        # 1. 統一將空字串、None 替換為 pandas 的標準空值 (NaN)
                        df.replace(r'^\s*$', pd.NA, regex=True, inplace=True)
                        
                        # 2. 刪除整欄都是空值的欄位（收緊表格）
                        df.dropna(how='all', axis=1, inplace=True)
                        
                        # ======== 新增：核心「向右對齊」邏輯 ========
                        # 建立一個函數，處理單一橫列：將所有 NaN 推到左邊，非 NaN 放到右邊
                        def shift_row_right(row):
                            # 取出這行所有不是 NaN 的值
                            valid_values = row.dropna().tolist()
                            # 計算需要補幾個 NaN 在最左邊
                            num_nans_to_add = len(row) - len(valid_values)
                            # 組合：[NaN, NaN...] + [值1, 值2...]
                            shifted_values = [pd.NA] * num_nans_to_add + valid_values
                            return pd.Series(shifted_values, index=row.index)
                        
                        # 將此邏輯套用到 DataFrame 的每一列 (axis=1)
                        df = df.apply(shift_row_right, axis=1)
                        # =======================================
                        
                        safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', section_name)[:31]
                        
                        # 寫入 Excel (保留表頭，移除欄位索引)
                        df.to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)
                        has_data = True

        if has_data:
            st.success("✅ 智能轉換成功！所有數據已自動向右對齊。請點擊下方按鈕下載 Excel。")
            st.download_button(
                label="📥 下載向右對齊 Excel 檔案",
                data=output.getvalue(),
                file_name="DSE_Report_Right_Aligned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("⚠️ 未能從上傳的 PDF 中找到有效的成績表格數據。")
