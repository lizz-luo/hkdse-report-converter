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
                # 尋找題號
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
                
                # 尋找選項與數據 (\uf0fe 是正確答案打勾符號)
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
            
            # 保存最後一題
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
# 使用者介面 (UI) 設計
# ==========================================
st.sidebar.header("⚙️ 設定選項")
report_type = st.sidebar.radio(
    "請選擇你要轉換的報告類型：",
    ("📝 項目分析報告 (Item Analysis)", "✅ 多項選擇題報告 (MCQ Analysis)")
)

st.markdown("請在下方上傳考評局的 DSE PDF 報告，系統會根據左側選擇的模式進行轉換。")

uploaded_file = st.file_uploader("拖曳或選擇 PDF 檔案", type=["pdf"])

if uploaded_file is not None:
    st.info("檔案讀取中，請稍候...")
    
    try:
        # 根據選擇的模式調用不同的處理函數
        if "Item Analysis" in report_type:
            df = extract_item_analysis(uploaded_file)
            sheet_name = "Item Analysis"
            export_name = "項目分析報告"
        else:
            df = extract_mcq_analysis(uploaded_file)
            sheet_name = "MCQ Analysis"
            export_name = "MCQ分析報告"
        
        # 顯示結果
        if df.empty:
            st.warning(f"⚠️ 無法提取數據！請確認你上傳的 PDF 是否為**{export_name}**，且模式選擇正確。")
        else:
            st.success(f"✅ 成功提取 {len(df)} 筆數據！")
            
            st.subheader("數據預覽")
            st.dataframe(df, use_container_width=True)
            
            # 下載按鈕
            excel_data = convert_df_to_excel(df, sheet_name)
            file_base_name = uploaded_file.name.replace('.pdf', '')
            
            st.download_button(
                label=f"📥 下載 Excel 檔案 ({export_name})",
                data=excel_data,
                file_name=f"{file_base_name}_{export_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
    except Exception as e:
        st.error(f"❌ 處理檔案時發生錯誤：{str(e)}")
