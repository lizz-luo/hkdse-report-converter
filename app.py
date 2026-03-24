import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

# ==========================================
# 頁面設定 / Page Configuration
# ==========================================
st.set_page_config(page_title="HKDSE Statistical Report Data Converter | HKDSE學校統計報告 數據轉換工具", page_icon="🔁", layout="wide")

st.title("📊 HKDSE學校統計報告 數據轉換工具")
st.markdown("本工具將自動提取考評局 PDF 報告中的數據，轉換為 Excel 格式，方便貼上至 CUHK QSIP 分析工具。")

# ==========================================
# 側邊欄：共用上傳區 / Sidebar: Global Upload
# ==========================================
with st.sidebar:
    st.header("📂 檔案上載區")
    st.info("💡 你可以在此上載包含完整數據的考評局 PDF，所有標籤頁將自動共用此檔案進行分析。")
    global_file = st.file_uploader("上載 PDF 報告 | Upload PDF", type=["pdf"], key="global_file")
    st.caption("🛡️ 本工具僅在記憶體中暫存 PDF，處理後立即刪除，不會儲存至硬碟或雲端。")


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
        "Item", "Max Mark", "Your school Attm. No.", 
        "Your school Attem.  %", "Your school Mean", "Your school Mean %", 
        "Your school SD", "Day schools Attem.  %", "Day schools Mean", 
        "Day schools Mean %", "Day schools SD"
    ]
    df = pd.DataFrame(extracted_data, columns=columns)
    
    numeric_cols = [
        "Max Mark", "Your school Attm. No.",
        "Your school Attem.  %", "Your school Mean", "Your school SD", 
        "Day schools Attem.  %", "Day schools Mean", "Day schools SD"
    ]
    
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    pct_cols = ["Your school Mean %", "Day schools Mean %"]
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
# 核心處理函數 3：總數分析及圖表 (Total Analysis)
# ==========================================
import plotly.graph_objects as go

@st.cache_data
def extract_latest_dse_total_data(file_bytes):
    target_grades = ['5**', '5*+', '5+', '4+', '3+', '2+', '1+', 'UNCL', '出席 Sat']
    results = []

    with pdfplumber.open(file_bytes) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            if "總數" in text and "貴校" in text and "5**" in text:
                lines = text.split('\n')
                in_total_section = False

                for line in lines:
                    if "總數 Total" in line or "總數" in line:
                        in_total_section = True
                    elif "男生 Male" in line or "女生 Female" in line:
                        in_total_section = False

                    if in_total_section:
                        clean_line = line.replace(',', '')
                        for grade in target_grades:
                            if clean_line.startswith(grade + " "):
                                parts = clean_line.split(grade)

                                if len(parts) >= 3:
                                    ys_numbers = parts[1].strip().split()
                                    ds_numbers = parts[2].strip().split()

                                    if ys_numbers and ds_numbers:
                                        if not any(r['等級'] == grade for r in results):
                                            results.append({
                                                '等級': grade,
                                                '貴校': int(ys_numbers[-1]),
                                                '日校': int(ds_numbers[-1])
                                            })
                                break

                if len(results) == len(target_grades):
                    break

    df = pd.DataFrame(results)
    if not df.empty:
        df['等級'] = pd.Categorical(df['等級'], categories=target_grades, ordered=True)
        df = df.sort_values('等級').reset_index(drop=True)

    return df

def generate_dse_charts(df_raw):
    # 計算單獨等級
    levels_single = ['5**', '5*', '5', '4', '3', '2', '1', 'UNCL']

    # 貴校單獨等級人數 (已反轉，5** 在前)
    ys_single_nums = [
        df_raw.loc[df_raw['等級']=='5**', '貴校'].values[0],
        df_raw.loc[df_raw['等級']=='5*+', '貴校'].values[0] - df_raw.loc[df_raw['等級']=='5**', '貴校'].values[0],
        df_raw.loc[df_raw['等級']=='5+', '貴校'].values[0] - df_raw.loc[df_raw['等級']=='5*+', '貴校'].values[0],
        df_raw.loc[df_raw['等級']=='4+', '貴校'].values[0] - df_raw.loc[df_raw['等級']=='5+', '貴校'].values[0],
        df_raw.loc[df_raw['等級']=='3+', '貴校'].values[0] - df_raw.loc[df_raw['等級']=='4+', '貴校'].values[0],
        df_raw.loc[df_raw['等級']=='2+', '貴校'].values[0] - df_raw.loc[df_raw['等級']=='3+', '貴校'].values[0],
        df_raw.loc[df_raw['等級']=='1+', '貴校'].values[0] - df_raw.loc[df_raw['等級']=='2+', '貴校'].values[0],
        df_raw.loc[df_raw['等級']=='UNCL', '貴校'].values[0]
    ]

    # 日校單獨等級人數 (已反轉，5** 在前)
    ds_single_nums = [
        df_raw.loc[df_raw['等級']=='5**', '日校'].values[0],
        df_raw.loc[df_raw['等級']=='5*+', '日校'].values[0] - df_raw.loc[df_raw['等級']=='5**', '日校'].values[0],
        df_raw.loc[df_raw['等級']=='5+', '日校'].values[0] - df_raw.loc[df_raw['等級']=='5*+', '日校'].values[0],
        df_raw.loc[df_raw['等級']=='4+', '日校'].values[0] - df_raw.loc[df_raw['等級']=='5+', '日校'].values[0],
        df_raw.loc[df_raw['等級']=='3+', '日校'].values[0] - df_raw.loc[df_raw['等級']=='4+', '日校'].values[0],
        df_raw.loc[df_raw['等級']=='2+', '日校'].values[0] - df_raw.loc[df_raw['等級']=='3+', '日校'].values[0],
        df_raw.loc[df_raw['等級']=='1+', '日校'].values[0] - df_raw.loc[df_raw['等級']=='2+', '日校'].values[0],
        df_raw.loc[df_raw['等級']=='UNCL', '日校'].values[0]
    ]

    ys_total = df_raw.loc[df_raw['等級']=='出席 Sat', '貴校'].values[0]
    ds_total = df_raw.loc[df_raw['等級']=='出席 Sat', '日校'].values[0]

    ys_pct = [n / ys_total * 100 if ys_total > 0 else 0 for n in ys_single_nums]
    ds_pct = [n / ds_total * 100 if ds_total > 0 else 0 for n in ds_single_nums]

    # === 柱狀圖 ===
    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(
        x=levels_single, y=ys_pct, name='Your school',
        marker_color='#99C2E6', text=[f"{p:.1f}%" for p in ys_pct], textposition='outside', cliponaxis=False
    ))
    fig_bar.add_trace(go.Bar(
        x=levels_single, y=ds_pct, name='Day schools',
        marker_color='#ED7D31', text=[f"{p:.1f}%" for p in ds_pct], textposition='outside', cliponaxis=False
    ))

    max_y = max(max(ys_pct), max(ds_pct))
    y_range = [0, max_y * 1.2 if max_y > 0 else 40]

    fig_bar.update_layout(
        title=dict(text="Comparison of Your school and Day schools - Bar chart", x=0.5, xanchor='center'),
        barmode='group',
        legend=dict(orientation='h', yanchor='bottom', y=-0.2, xanchor='center', x=0.5),
        margin=dict(l=40, r=20, t=60, b=60),
        plot_bgcolor='white',
    )
    fig_bar.update_yaxes(showgrid=True, gridwidth=1, gridcolor='lightgray', ticksuffix="%", range=y_range)
    fig_bar.update_xaxes(type='category')

    # === 折線圖 ===
    fig_line = go.Figure()
    fig_line.add_trace(go.Scatter(
        x=levels_single, y=ys_pct, name='Your school', mode='lines', line=dict(color='#5B9BD5', width=3)
    ))
    fig_line.add_trace(go.Scatter(
        x=levels_single, y=ds_pct, name='Day schools', mode='lines', line=dict(color='#ED7D31', width=3)
    ))
    fig_line.update_layout(
        title=dict(text="Comparison of Your school and Day schools - Line graph", x=0.5, xanchor='center'),
        legend=dict(orientation='h', yanchor='bottom', y=-0.2, xanchor='center', x=0.5),
        margin=dict(l=40, r=20, t=60, b=60),
        plot_bgcolor='white',
    )
    fig_line.update_xaxes(type='category')
    fig_line.update_yaxes(showgrid=True, gridwidth=1, gridcolor='lightgray', ticksuffix="%", range=[0, 40], dtick=5)

    return fig_bar, fig_line


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
tab0, tab1, tab2 = st.tabs(["📊 總數及圖表 Total Analysis & Charts", "📝 項目分析報告 Item Analysis Report", "✅ 多項選擇題報告 MCQ Analysis Report"])

# -----------------
# 標籤頁 0 的內容 / Tab 0 Content (總數分析及圖表)
# -----------------
with tab0:
    st.subheader("📊 總數及圖表轉換 | Total Analysis & Charts Converter")

    col_t1, col_t2 = st.columns([2, 5])
    with col_t1:
        st.info("""
        💡 **本區功能：**
        自動提取最新年份的「總數」數據，並直接生成對比柱狀圖與折線圖。

        **Function:**
        Automatically extracts the latest year's 'Total' data and generates comparative bar and line charts.
        """)
    with col_t2:
        if global_file is None:
            st.warning("👈 請先在左側邊欄上載 PDF 檔案 | Please upload a PDF file in the sidebar first.")
        else:
            with st.spinner("系統正在提取最新年份數據並繪製圖表... | Processing and rendering charts..."):
                try:
                    # Reset pointer just in case
                    global_file.seek(0)
                    df_total = extract_latest_dse_total_data(global_file)
                    if df_total.empty:
                        st.error("❌ 無法提取數據！請確認你上載的 PDF 包含「總數」表格。")
                    else:
                        st.success(f"✅ 提取成功！已取得最新年份數據。")

                        # 顯示數據表
                        st.subheader("📋 最新年份數據概覽 | Latest Year Data")
                        st.dataframe(df_total, use_container_width=True)

                        # 顯示圖表
                        st.subheader("📈 表現對比圖表 | Performance Charts")
                        fig_bar, fig_line = generate_dse_charts(df_total)

                        st.plotly_chart(fig_bar, use_container_width=True)
                        st.plotly_chart(fig_line, use_container_width=True)

                        st.download_button(
                            label="📥 下載數據 Excel | Download Data (Excel)",
                            data=convert_df_to_excel(df_total, "Total Analysis"),
                            file_name=f"{global_file.name.replace('.pdf', '')}_TotalData.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_total",
                            type="primary"
                        )
                except Exception as e:
                    st.error(f"❌ 處理檔案時發生錯誤：{str(e)}")



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
        if global_file is None:
            st.warning("👈 請先在左側邊欄上載 PDF 檔案 | Please upload a PDF file in the sidebar first.")
        else:
            with st.spinner("系統正在處理檔案，請稍候... | Processing file, please wait..."):
                try:
                    global_file.seek(0)
                    df_item = extract_item_analysis(global_file)
                    if df_item.empty:
                        st.error("❌ 無法提取數據！請確認你上載的是否為正確的「項目分析報告」。 \n *Failed to extract data! Please ensure you uploaded the correct 'Item Analysis Report'.*")
                    else:
                        st.success(f"✅ 提取成功！共獲取 {len(df_item)} 行數據。 \n *Extraction successful! {len(df_item)} rows retrieved.*")
                        
                        st.subheader("📋 數據概覽 | Data Preview")
                        st.dataframe(df_item, use_container_width=True)
                        
                        st.download_button(
                            label="📥 下載 Excel 檔案 | Download Excel File",
                            data=convert_df_to_excel(df_item, "Item Analysis"),
                            file_name=f"{global_file.name.replace('.pdf', '')}_ItemAnalysis.xlsx",
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
        st.caption("🛡️ 本工具僅在記憶體中暫存 PDF，處理後立即刪除，不會儲存至硬碟或雲端。 | PDFs are held temporarily in RAM only and deleted after processing. No storage on disk or cloud.")

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
                            file_name=f"{global_file.name.replace('.pdf', '')}_MCQAnalysis.xlsx",
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
