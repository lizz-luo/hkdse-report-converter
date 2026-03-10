import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="DSE 報告轉換工具", page_icon="📊", layout="wide")

# ==========================================
# 大標題 + 總覽說明
# ==========================================
st.title("📊 DSE 考評局報告轉換工具")
st.markdown("**專為學校老師設計，3 步驟轉換考評局 PDF → QSIP Excel 可用數據！**")

# ==========================================
# 總覽步驟說明 + 圖片
# ==========================================
col1, col2, col3 = st.columns(3)
with col1:
    st.info("**📋 步驟 1**")
    st.markdown("""
    1. 點擊上方標籤選擇報告類型
    2. 上傳考評局的 PDF 檔案
    3. 等待「✅ 成功提取」訊息
    """)
with col2:
    if st.image("step1_upload.png", use_column_width=True):
        pass
with col3:
    st.info("**📋 步驟 2**")
    st.markdown("""
    1. 點擊「📥 下載 Excel」
    2. 打開你內部 QSIP 分析檔案
    3. **複製對應工作表** → **貼上**！
    """)
    if st.image("step2_paste_qsip.png", use_column_width=True):
        pass

st.divider()

# ==========================================
# 核心處理函數 (與前版相同，略過顯示以節省篇幅)
# ==========================================
@st.cache_data
def extract_item_analysis(file_bytes):
    # ... (與前版完全相同)
    pass

@st.cache_data
def extract_mcq_analysis(file_bytes):
    # ... (與前版完全相同)
    pass

def convert_df_to_excel(df, sheet_name):
    # ... (與前版完全相同)
    pass

# ==========================================
# 標籤頁 1：項目分析報告
# ==========================================
tab1, tab2 = st.tabs(["📝 項目分析報告", "✅ 多項選擇題報告"])

with tab1:
    st.markdown("### 📝 項目分析報告轉換區")
    
    # 專屬說明區塊
    col_ex1, col_ex2 = st.columns([1, 2])
    with col_ex1:
        st.image("step1_item_report.png", caption="這就是你要上傳的項目分析報告", use_column_width=True)
    with col_ex2:
        st.info("**適用檔案：**")
        st.markdown("""
        ✅ 數學、英文等科目的「項目分析」報告
        ✅ 包含「平均分」「標準差」「差距 %」的表格
        ✅ **已自動移除兩個無用差距欄位**
        """)
    
    # 上傳區
    file_item = st.file_uploader("👆 請上傳項目分析 PDF", type=["pdf"], key="file_item")
    
    if file_item is not None:
        with st.spinner("正在分析 PDF，請稍候..."):
            df_item = extract_item_analysis(file_item)
            
        if df_item.empty:
            st.error("❌ 無法提取數據！請確認是否為項目分析報告。")
        else:
            st.success(f"✅ 成功提取 **{len(df_item)} 筆數據**！")
            
            # 數據預覽
            st.subheader("📋 數據預覽 (已轉數字格式)")
            st.dataframe(df_item, use_container_width=True)
            
            # QSIP 貼上說明
            st.info("**🎯 貼到 QSIP Excel：**")
            st.markdown("""
            1. 點擊「📥 下載項目分析 Excel」
            2. 打開你的 QSIP 檔案 → 找到「項目分析」工作表
            3. **全選** → **貼上** (Ctrl+V)
            4. ✅ 數字已完美對齊，可直接分析！
            """)
            
            st.download_button(
                label="📥 下載項目分析 Excel",
                data=convert_df_to_excel(df_item, "Item Analysis"),
                file_name=f"{file_item.name.replace('.pdf', '')}_項目分析.xlsx",
                type="primary"
            )

with tab2:
    st.markdown("### ✅ 多項選擇題報告轉換區")
    
    # 專屬說明區塊
    col_ex1, col_ex2 = st.columns([1, 2])
    with col_ex1:
        st.image("step2_mcq_report.png", caption="這就是你要上傳的 MCQ 報告", use_column_width=True)
    with col_ex2:
        st.info("**適用檔案：**")
        st.markdown("""
        ✅ 英文科「多項選擇題分析」報告
        ✅ 包含 ABCD 選項人數 + 正確答案打勾的表格
        ✅ **自動識別正確答案**，數字已轉整數格式
        """)
    
    # 上傳區
    file_mcq = st.file_uploader("👆 請上傳 MCQ 分析 PDF", type=["pdf"], key="file_mcq")
    
    if file_mcq is not None:
        with st.spinner("正在分析 PDF，請稍候..."):
            df_mcq = extract_mcq_analysis(file_mcq)
            
        if df_mcq.empty:
            st.error("❌ 無法提取數據！請確認是否為 MCQ 報告。")
        else:
            st.success(f"✅ 成功提取 **{len(df_mcq)} 題**！")
            
            # 數據預覽
            st.subheader("📋 數據預覽 (人數已轉整數)")
            st.dataframe(df_mcq, use_container_width=True)
            
            # QSIP 貼上說明
            st.info("**🎯 貼到 QSIP Excel：**")
            st.markdown("""
            1. 點擊「📥 下載 MCQ 分析 Excel」
            2. 打開你的 QSIP 檔案 → 找到「MCQ 分析」工作表
            3. **全選** → **貼上** (Ctrl+V)
            4. ✅ 可直接計算答對率、畫圓餅圖！
            """)
            
            st.download_button(
                label="📥 下載 MCQ 分析 Excel",
                data=convert_df_to_excel(df_mcq, "MCQ Analysis"),
                file_name=f"{file_mcq.name.replace('.pdf', '')}_MCQ分析.xlsx",
                type="primary"
            )

# ==========================================
# 底部成功說明 + 聯絡資訊
# ==========================================
st.divider()
st.markdown("---")
st.success("**🎉 完成！** 你的數據已準備好貼到 QSIP Excel 進行深入分析！")
st.caption("❓ 如有問題請聯絡數據分析組 | 版本 2.1 | 支援數學/英文項目分析 & MCQ")
