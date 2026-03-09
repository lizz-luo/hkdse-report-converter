import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ==========================================
# 頁面設定
# ==========================================
st.set_page_config(page_title="DSE 報告轉換工具", page_icon="📊", layout="wide")

st.title("📊 DSE 考評局報告 PDF 轉 Excel 工具")
st.markdown("請選擇你要轉換的報告類型，並上傳對應的 PDF 檔案。")

# ==========================================
# 核心處理函數 1：項目分析報告 (Item Analysis)
# ==========================================
@st.cache_data
def extract_item_analysis(file_bytes):
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
                clean_line = " ".join(line.split())
                match = row_pattern.search(clean_line)
                if match:
                    extracted_data.append(match.groups())

    columns = [
        "項目/題號 (Item & Ref)", "滿分 (Max Mark)", "人數 No.", 
        "作答 Attempted % (貴校)", "平均分 Mean (貴校)", "平均分 Mean % (貴校)", 
        "標準差 S.D. (貴校)", "作答 Attempted % (日校)", "平均分 Mean (日校)", 
        "平均分 Mean % (日校)", "標準差 S.D. (日校)", "差距 Diff (b)-(c)", "差距 Diff %"
    ]
    return pd.DataFrame(extracted_data, columns=columns)

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
                            row[f'Your school {opt}_No.'] = question_answers.get(f'{opt}_your', '')
                            row[f'Day schools {opt}_No.'] = question_answers.get(f'{opt}_day', '')
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
                    row[f'Your school {opt}_No.'] = question_answers.get(f'{opt}_your', '')
                    row[f'Day schools {opt}_No.'] = question_answers.get(f'{opt}_day', '')
                mcq_data.append(row)

    df = pd.DataFrame(mcq_data)
    if not df.empty:
        column_order = [
            'Question Number', 'Corr. Ans',
            'Your school A_No.', 'Your school B_No.', 'Your school C_No.', 'Your school D_No.',
            'Day schools A_No.', 'Day schools B_No.', 'Day schools C_No.', 'Day schools D_No.'
        ]
        df = df[column_order]
    
    return df

# ==========================================
# 輔助函數：匯出 Excel
# ==========================================
def convert_df_to_excel(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# ==========================================
# 建立主畫面兩個標籤頁 (Tabs) 入口
# ==========================================
tab1, tab2 = st.tabs(["📝 項目分析報告 (Item Analysis)", "✅ 多項選擇題報告 (MCQ Analysis)"])

# -----------------
# 標籤頁 1 的內容
# -----------------
with tab1:
    st.subheader("項目分析報告轉換區")
    st.info("💡 請上傳包含「平均分、標準差、差距」的常規項目分析報告 PDF。")
    file_item = st.file_uploader("選擇項目分析 PDF", type=["pdf"], key="file_item")

    if file_item is not None:
        try:
            df_item = extract_item_analysis(file_item)
            if df_item.empty:
                st.warning("⚠️ 無法提取數據！請確認檔案格式是否正確。")
            else:
                st.success(f"✅ 成功提取 {len(df_item)} 筆數據！")
                st.dataframe(df_item, use_container_width=True)
                
                st.download_button(
                    label="📥 下載項目分析 Excel",
                    data=convert_df_to_excel(df_item, "Item Analysis"),
                    file_name=f"{file_item.name.replace('.pdf', '')}_項目分析.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_item",
                    type="primary"
                )
        except Exception as e:
            st.error(f"❌ 處理錯誤：{str(e)}")

# -----------------
# 標籤頁 2 的內容
# -----------------
with tab2:
    st.subheader("多項選擇題報告轉換區")
    st.info("💡 請上傳包含「ABCD選項、人數百分比、打勾正確答案」的 MCQ 報告 PDF。")
    file_mcq = st.file_uploader("選擇 MCQ 分析 PDF", type=["pdf"], key="file_mcq")

    if file_mcq is not None:
        try:
            df_mcq = extract_mcq_analysis(file_mcq)
            if df_mcq.empty:
                st.warning("⚠️ 無法提取數據！請確認檔案是否為 MCQ 報告。")
            else:
                st.success(f"✅ 成功提取 {len(df_mcq)} 題的數據！")
                st.dataframe(df_mcq, use_container_width=True)
                
                st.download_button(
                    label="📥 下載 MCQ 分析 Excel",
                    data=convert_df_to_excel(df_mcq, "MCQ Analysis"),
                    file_name=f"{file_mcq.name.replace('.pdf', '')}_MCQ分析.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_mcq",
                    type="primary"
                )
        except Exception as e:
            st.error(f"❌ 處理錯誤：{str(e)}")
