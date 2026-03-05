import streamlit as st
import pdfplumber
import pandas as pd
import io

# 設定網頁標題與排版
st.set_page_config(page_title="HKDSE 數據轉換工具", layout="centered")

# 網頁上的標題與說明文字
st.title("📊 HKDSE 項目分析報告轉 Excel 工具")
st.write("請上傳考評局的 PDF 報告（支援數學科：必修、M1、M2），系統會自動提取表格並轉換為 Excel。")

# 建立一個檔案上傳區塊
uploaded_file = st.file_uploader("請點擊或拖曳上傳 PDF 檔案", type="pdf")

if uploaded_file is not None:
    # 當使用者上傳檔案後，顯示載入中的動畫
    with st.spinner("正在解析 PDF 並提取數據，這可能需要幾秒鐘..."):
        # 準備空抽屜
        sections = {
            "Compulsory_Part": [],
            "M1_Calculus_Stats": [],
            "M2_Algebra_Calculus": []
        }
        current_section = None

        # 讀取上傳的 PDF
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                
                # 尋找科目標記
                if "數學必修部分" in text or "Mathematics Compulsory Part" in text:
                    current_section = "Compulsory_Part"
                elif "微積分與統計" in text or "Calculus and Statistics" in text:
                    current_section = "M1_Calculus_Stats"
                elif "代數與微積分" in text or "Algebra and Calculus" in text:
                    current_section = "M2_Algebra_Calculus"
                    
                # 提取表格
                table = page.extract_table()
                if table and current_section:
                    sections[current_section].extend(table)
                else:
                    table_alt = page.extract_table({"vertical_strategy": "text", "horizontal_strategy": "text"})
                    if table_alt and current_section:
                        sections[current_section].extend(table_alt)

        # 準備將數據寫入記憶體中的 Excel（不用存到實體硬碟）
        output = io.BytesIO()
        has_data = False
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for section_name, data in sections.items():
                if data:
                    # 數據清洗：移除空白行
                    clean_data = [row for row in data if any(cell.strip() for cell in row if isinstance(cell, str))]
                    if clean_data:
                        df = pd.DataFrame(clean_data)
                        df.to_excel(writer, sheet_name=section_name, index=False, header=False)
                        has_data = True

        # 顯示下載按鈕
        if has_data:
            st.success("✅ 轉換成功！請點擊下方按鈕下載 Excel。")
            st.download_button(
                label="📥 下載 Excel 檔案",
                data=output.getvalue(),
                file_name="DSE_Maths_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("⚠️ 未能從上傳的 PDF 中找到有效的成績表格數據，請確認檔案是否正確。")
