import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# 設定網頁標題與排版
st.set_page_config(page_title="DSE 報告轉換器", page_icon="📊", layout="wide")

st.title("📊 DSE 項目分析報告 PDF 轉 Excel 工具")
st.markdown("請上傳考評局的 DSE 數學科項目分析報告 (PDF)，系統會自動提取表格並轉換為 Excel 格式。")

@st.cache_data
def extract_dse_data(file_bytes):
    # 更新為貼合 pdfplumber 水平掃描輸出的正則表達式
    row_pattern = re.compile(
        r'^(.*?)\s+(\d+)\s+(\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+%)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+%)\s+(\d+\.\d+)\s*([+-]?\d+\.\d+)\s*([+-]?\d+%)$'
    )

    extracted_data = []
    
    with pdfplumber.open(file_bytes) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
                
            for line in text.split('\n'):
                # 統整空白，避免多餘空格干擾
                clean_line = " ".join(line.split())
                match = row_pattern.search(clean_line)
                
                if match:
                    extracted_data.append(match.groups())

    # 配合新順序更新欄位名稱 (這一次連項目和題號都智慧合併了)
    columns = [
        "項目/題號 (Item & Ref)", 
        "滿分 (Max Mark)", 
        "人數 No.", 
        "作答 Attempted % (貴校)", 
        "平均分 Mean (貴校)", 
        "平均分 Mean % (貴校)", 
        "標準差 S.D. (貴校)", 
        "作答 Attempted % (日校)", 
        "平均分 Mean (日校)", 
        "平均分 Mean % (日校)", 
        "標準差 S.D. (日校)", 
        "差距 Diff (b)-(c)", 
        "差距 Diff %"
    ]

    return pd.DataFrame(extracted_data, columns=columns)

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='DSE Data')
    processed_data = output.getvalue()
    return processed_data

uploaded_file = st.file_uploader("拖曳或選擇 PDF 檔案", type=["pdf"])

if uploaded_file is not None:
    st.info("檔案讀取中，請稍候...")
    
    try:
        df = extract_dse_data(uploaded_file)
        
        if df.empty:
            st.warning("⚠️ 無法從此 PDF 中提取到符合格式的數據，請確認檔案是否為標準的 DSE 項目分析報告。")
        else:
            st.success(f"✅ 成功提取 {len(df)} 筆數據！")
            
            st.subheader("數據預覽")
            st.dataframe(df, use_container_width=True)
            
            excel_data = convert_df_to_excel(df)
            
            st.download_button(
                label="📥 下載 Excel 檔案",
                data=excel_data,
                file_name=f"{uploaded_file.name.replace('.pdf', '')}_解析結果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
    except Exception as e:
        st.error(f"❌ 處理檔案時發生錯誤：{str(e)}")
