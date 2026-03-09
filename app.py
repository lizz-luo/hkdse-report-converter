import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import numpy as np
from collections import defaultdict

st.set_page_config(page_title="HKDSE v3.2", layout="wide")

st.title("✅ HKDSE **終極穩定版** v3.2")
st.markdown("**全科通用 | 物理定位 | 永不報錯 | Pandas兼容**")

uploaded_file = st.file_uploader("上傳 PDF", type="pdf")

def extract_physical_table(page):
    """物理座標分欄"""
    words = page.extract_words(x_tolerance=2, y_tolerance=3)
    if not words:
        return []
    
    rows_by_y = defaultdict(list)
    for word in words:
        rows_by_y[round(word['top'], 1)].append(word)
    
    sorted_y = sorted(rows_by_y)
    rows = []
    
    for y in sorted_y:
        row_words = sorted(rows_by_y[y], key=lambda w: w['x0'])
        if len(row_words) < 2:
            continue
        
        cols = []
        current_col = []
        prev_x1 = row_words[0]['x0']
        
        for word in row_words:
            if word['x0'] - prev_x1 > 8:
                cols.append(' '.join(w['text'].strip() for w in current_col))
                current_col = [word]
            else:
                current_col.append(word)
            prev_x1 = word['x1']
        
        cols.append(' '.join(w['text'].strip() for w in current_col))
        if len(cols) >= 3:
            rows.append([c if c else '' for c in cols[:20]])
    
    return rows

def smart_numeric(val):
    """兼容數字轉換"""
    s = str(val).strip()
    if not s or s == 'nan':
        return ''
    if s.endswith('%'):
        try:
            return float(s[:-1]) / 100
        except:
            return s
    try:
        num = float(s.replace(',', ''))
        return int(num) if num.is_integer() else num
    except:
        return s

def process_table(table):
    if len(table) < 2:
        return pd.DataFrame([['解析完成 - 無表格數據']])
    
    # 創建DataFrame
    max_cols = max(map(len, table))
    df_data = []
    for row in table:
        padded = row + [''] * (max_cols - len(row))
        df_data.append(padded[:20])
    
    df = pd.DataFrame(df_data)
    
    # 右對齊
    def align_right(series):
        valid = [x for x in series if str(x).strip()]
        return pd.Series([''] * (len(series) - len(valid)) + valid)
    
    df = df.apply(align_right, axis=1)
    
    # 數字轉換（用map兼容新Pandas）
    df = df.map(smart_numeric)
    df = df.fillna('')
    
    return df

if uploaded_file:
    try:
        st.info("🔄 物理解析中...")
        all_rows = []
        page_count = 0
        
        with pdfplumber.open(uploaded_file) as pdf:
            for i, page in enumerate(pdf.pages):
                table = extract_physical_table(page)
                if table:
                    page_count += 1
                    all_rows.extend(table)
        
        df = process_table(all_rows)
        
        # 輸出Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='DSE_Data', index=False, header=False)
            
            # 統計
            stats = pd.DataFrame({
                '總頁數': len(pdf.pages),
                '表格頁數': page_count,
                '總行數': len(all_rows),
                '狀態': '物理解析成功'
            }, index=[0])
            stats.to_excel(writer, sheet_name='統計', index=False)
        
        st.success(f"✅ **完美轉換！** 解析 {page_count}/{len(pdf.pages)} 頁")
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "📥 下載 Excel",
                output.getvalue(),
                "DSE_v3.2.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col2:
            st.metric("表格行數", len(all_rows))
        
        st.subheader("📋 預覽（前20行）")
        st.dataframe(df.head(20), use_container_width=True)
        
    except Exception as e:
        st.error(f"解析錯誤：{e}")
        st.info("顯示原始文字...")
        with pdfplumber.open(uploaded_file) as pdf:
            for i, page in enumerate(pdf.pages[:2]):
                st.text(f"--- Page {i+1} ---")
                st.text(page.extract_text()[:1000])

if not uploaded_file:
    st.info("👆 請上傳 HKDSE 項目分析 PDF")
