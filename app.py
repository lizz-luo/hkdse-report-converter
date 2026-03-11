import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

# ==========================================
# 頁面設定 / Page Configuration
# ==========================================
st.set_page_config(page_title="HKDSE Stastical Report Data Converter | HKDSE學校統計報告 數據轉換工具", page_icon="🔁", layout="wide")

st.title("HKDSE學校統計報告 數據轉換工具 | HKDSE Stastical Report Data Converter")
st.markdown("""
請選擇你要轉換的報告類型，並上載相關的 PDF 檔案。 本工具將自動提取有用數據，並轉換為 Excel 格式，以便貼上至 QSIP 分析工具。 \n\n

*Please select the report type and upload the corresponding PDF file. This tool will extract useful data and convert it into Excel format that is ready to be pasted into the QSIP analysis tool.*
""")

# ==========================================
# 核心處理函數 1：項目分析報告 (Item Analysis)
# ==========================================
@st.cache_data
def extract_item_analysis(file_bytes):
    row_pattern = re.compile(
        r'^(.*?)\s+(\d+)\s+(\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+%)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+%)\s+(\d+\.\d+)\s*([+-]?\d+\.\d+)\s*'
    )
    extracted_data = []
    
    with pdfplumber.open(file_bytes) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
                
            for line in text.split('\n'):
                clean_line = " ".join(line.split())
                match = row_pattern.search(clean_line)
                if match:
                    extracted_data.append(match.groups()[:11])

    columns = [
        "項目/題號 (Item & Ref)", "滿分 (Max Mark)", "人數 No.", 
        "作答 Attempted % (貴校 Your Sch)", "平均分 Mean (貴校 Your Sch)", "平均分 Mean % (貴校 Your Sch)", 
        "標準差 S.D. (貴校 Your Sch)", "作答 Attempted % (日校 Day Sch)", "平均分 Mean (日校 Day Sch)", 
        "平均分 Mean % (日校 Day Sch)", "標準差 S.D. (日校 Day Sch)"
    ]
    df = pd.DataFrame(extracted_data, columns=columns)
    
    numeric_cols = [
        "滿分 (Max Mark)", "人數 No.",
        "作答 Attempted % (貴校 Your Sch)", "平均分 Mean (貴校 Your Sch)", "標準差 S.D. (貴校 Your Sch)", 
        "作答 Attempted % (日校 Day Sch)", "平均分 Mean (日校 Day Sch)", "標準差 S.D. (日校 Day Sch)"
    ]
    
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    pct_cols = ["平均分 Mean % (貴校 Your Sch)", "平均分 Mean % (日校 Day Sch)"]
    for col in pct_cols:
        df[col] = df[col].str.replace('%', '').astype(float) / 100
    
    return df

# ==========================================
# 核心處理函數 2：多項選擇題報告 (MCQ Analysis)
# ==========================================
@st.cache_data
def extract_mcq_analysis(file_bytes):
    mcq_data = []
    
    with pdfplumber.open(file_bytes) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            
            lines = text.split('\n')
            current_question = None
            correct_answer = None
            question_answers = {}
            
            for line in lines:
                q_match = re.match(r'^(\d+\([ivx]+\)|\d+)\s+貴校', line.strip())
                if q_match:
                    if current_question and question_answers:
                        row = {'Question Number': current_question, 'Corr. Ans': correct_answer}
                        for opt in ['A', 'B', 'C', 'D']:
                            row[f'Your school {opt}_No.'] = question_answers.get(f'{opt}_your', '0')
                            row[f'Day schools {opt}_No.'] = question_answers.get(f'{opt}_day', '0')
                        mcq_data.append(row)
                    
                    current_question = q_match.group(1)
                    question_answers = {}
                    correct_answer = None
                
                answer_match = re.match(r'^([ABCD])\s+(\uf0fe)?\s*(\d+)\s+[\d.]+\s+([\d,]+)', line.strip())
                if answer_match and current_question:
                    option = answer_match.group(1)
                    has_marker = answer_match.group(2) is not None
                    your_no = answer_match.group(3)
                    day_no = answer_match.group(4).replace(',', '')
                    
                    if has_marker:
                        correct_answer = option
                    
                    question_answers[f'{option}_your'] = your_no
                    question_answers[f'{option}_day'] = day_no
            
            if current_question and question_answers:
                row = {'Question Number': current_question, 'Corr. Ans': correct_answer}
                for opt in ['A', 'B', 'C', 'D']:
                    row[f'Your school {opt}_No.'] = question_answers.get(f'{opt}_your', '0')
                    row[f'Day schools {opt}_No.'] = question_answers.get(f'{opt}_day', '0')
                mcq_data.append(row)

    df = pd.DataFrame(mcq_data)
    if not df.empty:
        column_order = [
            'Question Number', 'Corr. Ans',
            'Your school A_No.', 'Your school B_No.', 'Your school C_No.', 'Your school D_No.',
            'Day schools A_No.', 'Day schools B_No.', 'Day schools C_No.', 'Day schools D_No.'
        ]
        df = df[column_order]
        
        for col in df.columns:
            if '_No.' in col:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    
    return df

# ==========================================
# 輔助函數：匯出 Excel / Export to Excel
# ==========================================
def convert_df_to_excel(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# ==========================================
# 建立主畫面兩個標籤頁 (Tabs) 入口
# ==========================================
tab1, tab2 = st.tabs(["📝 項目分析報告 Item Analysis Report", "✅ 多項選擇題報告 MCQ Analysis Report"])

# -----------------
# 標籤頁 1 的內容 / Tab 1 Content
# -----------------
with tab1:
    st.subheader("📝 項目分析報告轉換 | Item Analysis Converter")
    
    col1, col2 = st.columns([2, 5])
    
    with col1:
        st.info("""
        💡 **本區適用於以下格式的報告：**
        表格橫向列出「平均分 Mean」、「標準差 S.D.」等數據。
        
        **Applicable for reports formatted like:**
        The table horizontally displays data such as 'Mean' and 'S.D.'.
        """)
        if os.path.exists("example1_item.png"):
            st.image("example1_item.png", caption="項目分析表格示例 | Example of Item Analysis Table", use_column_width=True)
        else:
            st.warning("⚠️ (提示: 系統未找到 example1_item.png | Image not found)")
            
    with col2:
        file_item = st.file_uploader("📂 請於此處上載「項目分析」PDF  |  Upload 'Item Analysis' PDF here", type=["pdf"], key="file_item")

        if file_item is not None:
            with st.spinner("系統正在處理檔案，請稍候... | Processing file, please wait..."):
                try:
                    df_item = extract_item_analysis(file_item)
                    if df_item.empty:
                        st.error("❌ 無法提取數據！請確認你上載的是否為正確的「項目分析報告」。 \n *Failed to extract data! Please ensure you uploaded the correct 'Item Analysis Report'.*")
                    else:
                        st.success(f"✅ 提取成功！共獲取 {len(df_item)} 行數據。 \n *Extraction successful! {len(df_item)} rows retrieved.*")
                        
                        st.subheader("📋 數據概覽 | Data Preview")
                        st.dataframe(df_item, use_container_width=True)
                        
                        st.download_button(
                            label="📥 下載 Excel 檔案 | Download Excel File",
                            data=convert_df_to_excel(df_item, "Item Analysis"),
                            file_name=f"{file_item.name.replace('.pdf', '')}_ItemAnalysis.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_item",
                            type="primary"
                        )
                except Exception as e:
                    st.error(f"❌ 處理檔案時發生錯誤 | Error processing file：{str(e)}")

# -----------------
# 標籤頁 2 的內容 / Tab 2 Content
# -----------------
with tab2:
    st.subheader("✅ 多項選擇題報告轉換 | MCQ Analysis Converter")
    
    col3, col4 = st.columns([2, 5])
    
    with col3:
        st.info("""
        💡 **本區適用於以下格式的報告：**
        表格列出「A, B, C, D」選項的選擇人數，並附有 ☑️ 標記顯示正確答案。
        
        **Applicable for reports formatted like:**
        The table lists the number of students for options 'A, B, C, D' and uses a ☑️ mark to indicate the correct answer.
        """)
        if os.path.exists("example2_mcq.png"):
            st.image("example2_mcq.png", caption="多項選擇題表格示例 | Example of MCQ Analysis Table", use_column_width=True)
        else:
            st.warning("⚠️ (提示: 系統未找到 example2_mcq.png | Image not found)")
            
    with col4:
        file_mcq = st.file_uploader("📂 請於此處上載「多項選擇題分析」PDF  |  Upload 'MCQ Analysis' PDF here", type=["pdf"], key="file_mcq")

        if file_mcq is not None:
            with st.spinner("系統正在處理檔案，請稍候... | Processing file, please wait..."):
                try:
                    df_mcq = extract_mcq_analysis(file_mcq)
                    if df_mcq.empty:
                        st.error("❌ 無法提取數據！請確認你上載的是否為正確的「多項選擇題分析報告」。 \n *Failed to extract data! Please ensure you uploaded the correct 'MCQ Analysis Report'.*")
                    else:
                        st.success(f"✅ 提取成功！共獲取 {len(df_mcq)} 題的數據。 \n *Extraction successful! Data for {len(df_mcq)} questions retrieved. *")
                        
                        st.subheader("📋 數據概覽 | Data Preview")
                        st.dataframe(df_mcq, use_container_width=True)
                        
                        st.download_button(
                            label="📥 下載 Excel 檔案 | Download Excel File",
                            data=convert_df_to_excel(df_mcq, "MCQ Analysis"),
                            file_name=f"{file_mcq.name.replace('.pdf', '')}_MCQAnalysis.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_mcq",
                            type="primary"
                        )
                except Exception as e:
                    st.error(f"❌ 處理檔案時發生錯誤 | Error processing file：{str(e)}")

# ==========================================
# 頁尾提示 / Footer Notes
# ==========================================
st.divider()
st.caption("""
📌 **小貼士 Tips:** 
下載 Excel 後，請打開檔案，選中並複製(Ctrl+C)轉換結果，然後直接貼上(Ctrl+V)至QSIP HKDSE分析工具。 \n
*After downloading the Excel file, please open it, select and copy (Ctrl+C) the conversion results, and then paste (Ctrl+V) them directly into the QSIP HKDSE Analysis Tool.*
""")
