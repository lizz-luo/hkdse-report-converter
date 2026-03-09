import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# 設定網頁標題與排版
st.set_page_config(page_title="DSE 報告轉換器", page_icon="📊", layout="wide")

st.title("📊 DSE 項目分析報告 PDF 轉 Excel 工具")
st.markdown("請上傳考評局的 DSE 數學科項目分析報告 (PDF)，系統會自動提取表格並轉換為 Excel 格式。")

# 定義核心提取函數，並加入快取避免重複運算
@st.cache_data
def extract_dse_data(file_bytes):
    # DSE 數據行的正則表達式
    row_pattern = re.compile(
        r'^(\d+)\s+([Q\d\w\(\)]+)\s+(.+?)\s+([+-]?\d+%)\s+([+-]?\d+\.\d+)\s+(\d+\.\d+)\s+(\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+%)\s+(\d+%)\s+(\d+\.\d+)\s+(\d+\.\d+)$'
    )

    extracted_data = []
    
    # 使用 pdfplumber 讀取上傳的二進制檔案
    with pdfplumber.open(file_bytes) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
                
            for line in text.split('\n'):
                clean_line = " ".join(line.split())
                match = row_pattern.search(clean_line)
                
                if match:
                    extracted_data.append(match.groups())

    columns = [
        "滿分 (Max Mark)", "項目 (Item)", "試題/評卷編號 (Ref)", 
        "差距 Diff %", "差距 Diff (b)-(c)", "標準差 S.D. (貴校)", 
        "人數 No.", "作答 Attempted % (貴校)", "作答 Attempted % (日校)", 
        "平均分 Mean % (c)/(a)", "百分率 % (貴校)", "百分率 % (日校)", 
        "平均分 Mean (貴校)", "平均分 Mean (日校)"
    ]

    return pd.DataFrame(extracted_data, columns=columns)

# 將 DataFrame 轉換為 Excel 的二進制流，供下載使用
def convert_df_to_excel(df):
    output = io.BytesIO()
    # 使用 openpyxl 引擎寫入內存
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='DSE Data')
    processed_data = output.getvalue()
    return processed_data

# 建立上傳區塊
uploaded_file = st.file_uploader("拖曳或選擇 PDF 檔案", type=["pdf"])

if uploaded_file is not None:
    st.info("檔案讀取中，請稍候...")
    
    try:
        # 執行數據提取
        df = extract_dse_data(uploaded_file)
        
        if df.empty:
            st.warning("⚠️ 無法從此 PDF 中提取到符合格式的數據，請確認檔案是否為標準的 DSE 項目分析報告。")
        else:
            st.success(f"✅ 成功提取 {len(df)} 筆數據！")
            
            # 顯示數據預覽
            st.subheader("數據預覽")
            st.dataframe(df, use_container_width=True)
            
            # 準備下載按鈕
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
