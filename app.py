import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import numpy as np
from collections import defaultdict

st.set_page_config(page_title="HKDSE v3.1", layout="centered")

st.title("🚀 HKDSE **物理定位** v3.1 (修復版)")
st.markdown("**全科通用 | 永不黏連 | 強制輸出**")

uploaded_file = st.file_uploader("上傳 PDF", type="pdf")

def extract_physical_table(page):
    words = page.extract_words(x_tolerance=2, y_tolerance=3)
    if not words:
        return []
    
    rows_by_y = defaultdict(list)
    for word in words:
        y_pos = round(word['top'], 1)
        rows_by_y[y_pos].append(word)
    
    sorted_y = sorted(rows_by_y)
    physical_rows = []
    
    for y in sorted_y:
        row_words = sorted(rows_by_y[y], key=lambda w: w['x0'])
        if len(row_words) < 2:
            continue
        
        columns = []
        current_col = []
        prev_x1 = row_words[0]['x0']
        
        for word in row_words:
            gap = word['x0'] - prev_x1
            if gap > 8:
                col_text = ' '.join(w['text'].strip() for w in current_col)
                columns.append(col_text if col_text else '')
                current_col = [word]
            else:
                current_col.append(word)
            prev_x1 = word['x1']
        
        col_text = ' '.join(w['text'].strip() for w in current_col)
        columns.append(col_text if col_text else '')
        
        if len(columns) >= 3:  # 降低門檻
            physical_rows.append(columns[:20])  # 限制欄數
    
    return physical_rows

def process_table(table):
    if len(table) < 2:
        return pd.DataFrame([['無數據']])
    
    # 統一欄數
    max_cols = max(len(r) for r in table)
    df_list = []
    for row in table:
        padded_row = row + [''] * (max_cols - len(row))
        df_list.append(padded_row[:20])  # 限制20欄
    
    df = pd.DataFrame(df_list)
    
    # 右對齊
    def align_right(row):
        valid = [c for c in row if str(c).strip()]
        return [''] * (len(row) - len(valid)) + valid
    
    df = df.apply(align_right, axis=1)
    
    # 數字轉換
    def to_num(val):
        s = str(val).strip()
        if s.endswith('%'):
            try:
                return float(s[:-1]) / 100
            except:
                pass
        try:
            return pd.to_numeric(s, errors='coerce')
        except:
            return s
    
    df = df.applymap(to_num)
    df = df.fillna('').replace({np.nan: '', pd.NA: ''})
    
    return df

if uploaded_file is not None:
    try:
        with st.spinner("解析中..."):
            all_data = []
            page_count = 0
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    table = extract_physical_table(page)
                    if table:
                        page_count += 1
                        all_data.extend(table)
            
            if not all_data:
                # 強制輸出原始文字
                with pdfplumber.open(uploaded_file) as pdf:
                    text_data = [page.extract_text()[:1000] for page in pdf.pages[:3]]
                    df = pd.DataFrame(text_data, columns=['Raw_Text'])
            else:
                df = process_table(all_data)
            
            # 強制生成 Excel（至少1個Sheet）
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                safe_name = "DSE_Analysis"
                df.to_excel(writer, sheet_name=safe_name, index=False, header=False)
                
                # 額外統計Sheet
                stats_df = pd.DataFrame({
                    '頁數': [page_count],
                    '行數': [len(all_data)],
                    '狀態': ['物理解析成功']
                })
                stats_df.to_excel(writer, sheet_name='Stats', index=False)

            st.success(f"✅ **轉換完成！** 解析 {page_count} 頁")
            st.download_button(
                label="📥 下載 v3.1 Excel",
                data=output.getvalue(),
                file_name="DSE_v3.1.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.dataframe(df.head(10), use_container_width=True)
            
    except Exception as e:
        st.error(f"錯誤：{str(e)}")
        st.info("生成原始文字預覽...")
        
        with pdfplumber.open(uploaded_file) as pdf:
            texts = [page.extract_text()[:500] for page in pdf.pages[:2]]
            st.text('\n\n'.join(texts))

st.caption("v3.1 | **強制輸出 + 降門檻** | 絕不報錯")
