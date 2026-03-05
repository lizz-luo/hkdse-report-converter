import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="HKDSE 全科數據轉換工具", layout="centered")

st.title("📊 HKDSE 項目分析報告轉 Excel 工具")
st.write("請上傳考評局的 PDF 報告（支援**所有科目**，包括數學、英文等），系統會自動按「卷別 (Paper)」提取表格並轉換為 Excel。")

uploaded_file = st.file_uploader("請點擊或拖曳上傳 PDF 檔案", type="pdf")

if uploaded_file is not None:
    with st.spinner("正在智能解析 PDF 並提取各卷別數據，這可能需要幾秒鐘..."):
        # 使用字典來動態儲存不同 Paper 的數據，不再寫死科目
        sections = {}
        current_section = "General" # 預設名稱，以防第一頁找不到 Paper 名稱

        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                
                # 關鍵升級：使用正則表達式動態捕捉 "卷 Paper: XXX" 
                # 例如捕捉 "卷 Paper: 1A" 或 "卷 Paper: 2"
                paper_match = re.search(r'卷\s*Paper:\s*([A-Za-z0-9]+)', text)
                if paper_match:
                    # 抓取到的卷別名稱，例如 '1A', '1B1', 'M1'
                    paper_name = paper_match.group(1).strip()
                    # 確保工作表名稱不超過 Excel 限制（最多31字元）
                    current_section = f"Paper_{paper_name}"
                    
                    # 如果這個卷別還沒被建立過，就幫它準備一個新抽屜
                    if current_section not in sections:
                        sections[current_section] = []
                
                # 如果還沒抓到特定的 Paper，但有抓到科目總稱（確保容錯）
                elif "Mathematics Compulsory Part" in text and current_section == "General":
                    current_section = "Math_Compulsory"
                elif "English Language" in text and current_section == "General":
                    current_section = "Eng_Lang"
                
                # 如果字典裡還沒有這個 current_section，就初始化
                if current_section not in sections:
                    sections[current_section] = []
                    
                # 提取表格並放入對應的抽屜
                table = page.extract_table()
                if table:
                    sections[current_section].extend(table)
                else:
                    table_alt = page.extract_table({"vertical_strategy": "text", "horizontal_strategy": "text"})
                    if table_alt:
                        sections[current_section].extend(table_alt)

        output = io.BytesIO()
        has_data = False
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for section_name, data in sections.items():
                if data:
                    # 數據清洗：移除空白行
                    clean_data = [row for row in data if any(cell.strip() for cell in row if isinstance(cell, str))]
                    if clean_data:
                        # 將數據轉為 DataFrame
                        df = pd.DataFrame(clean_data)
                        
                        # 處理 Excel Sheet 名稱不能超過 31 個字元的限制，並替換掉不合法的字元
                        safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', section_name)[:31]
                        
                        df.to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)
                        has_data = True

        if has_data:
            st.success("✅ 智能轉換成功！系統已自動識別文件中的卷別。請點擊下方按鈕下載 Excel。")
            st.download_button(
                label="📥 下載全科 Excel 檔案",
                data=output.getvalue(),
                file_name="DSE_Report_Converted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("⚠️ 未能從上傳的 PDF 中找到有效的成績表格數據，請確認檔案格式是否為 HKEAA 的標準報告。")
