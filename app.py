import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import numpy as np

st.set_page_config(page_title="HKDSE 轉換工具", layout="wide")

st.title("HKDSE 轉換")

uploaded_file = st.file_uploader("上傳 PDF", type="pdf")

def smart_split_text(text):
    splits = re.split(r'[ /()#：:]+', text)
    return [s.strip() for s in splits if s.strip()]

def extract_smart_table(page):
    words = page.extract_words(x_tolerance=1.5, y_tolerance=2.5)
    if not words:
        return []
    
    rows_by_y = {}
    for word in words:
        y_pos = round(word['top'], 1)
        if y_pos not in rows_by_y:
            rows_by_y[y_pos] = []
        rows_by_y[y_pos].append(word)
    
    sorted_y = sorted(rows_by_y.keys())
    rows = []
    
    for y in sorted_y:
        row_words = sorted(rows_by_y[y], key=lambda w: w['x0'])
        if len(row_words) < 2:
            continue
        
        x_pos = [w['x0'] for w in row_words]
        gaps = [x_pos[i+1] - x_pos[i] for i in range(len(x_pos)-1)]
        dyn_gap = np.percentile([g for g in gaps if g > 0], 70)
        
        cols = []
        current_col = []
        prev_x1 = row_words[0]['x0']
        
        for word in row_words:
            gap = word['x0'] - prev_x1
            if gap > dyn_gap * 0.7:
                col_text = ' '.join(w['text'] for w in current_col).strip()
                if col_text:
                    cols.extend(smart_split_text(col_text))
                current_col = [word]
            else:
                current_col.append(word)
            prev_x1 = word['x1']
        
        col_text = ' '.join(w['text'] for w in current_col).strip()
        if col_text:
            cols.extend(smart_split_text(col_text))
        
        if len(cols) >= 4:
            rows.append([c if c else '' for c in cols[:25]])
    
    return rows

def align_to_rightmost(df):
    """對齊到最右有數據欄"""
    # 轉string避免str錯誤
    df_str = df.astype(str)
    
    # 找到最右有數據列
    has_data_cols = (df_str != '').any(axis=0)
    rightmost_col = has_data_cols[has_data_cols].index.max()
    
    # 統一寬度
    df = df.reindex(columns=range(rightmost_col + 1), fill_value='')
    
    # 右對齊
    def right_align_row(series):
        series_str = series.astype(str)
        valid_mask = series_str.str.strip() != ''
        valid = series[valid_mask].tolist()
        n_empty = len(series) - len(valid)
        return pd.Series([''] * n_empty + valid)
    
    df = df.apply(right_align_row, axis=1)
    
    # 刪左空欄
    df_str = df.astype(str)
    df = df.loc[:, (df_str != '').any(axis=0)]
    
    return df

def smart_numeric(val):
    s = str(val).strip()
    if s == 'None' or s == 'nan' or pd.isna(val):
        return ''
    if s.endswith('%'):
        try:
            return float(s[:-1]) / 100
        except:
            pass
    try:
        num = float(s.replace(',', ''))
        return int(num) if num.is_integer() else num
    except:
        return s

if uploaded_file:
    try:
        all_rows = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = extract_smart_table(page)
                all_rows.extend(table)
        
        df_raw = pd.DataFrame(all_rows)
        df_aligned = align_to_rightmost(df_raw)
        df_final = df_aligned.map(smart_numeric).fillna('').replace('None', '')
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, 'Data', index=False, header=False)
            
            stats = pd.DataFrame([{
                '行數': len(all_rows),
                '欄數': len(df_final.columns)
            }])
            stats.to_excel(writer, 'Stats', index=False)
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("下載 Excel", output.getvalue(), "DSE.xlsx")
        with col2:
            st.metric("行數", len(all_rows))
            st.metric("欄數", len(df_final.columns))
        
        st.subheader("數據預覽")
        st.dataframe(df_final.head(15), height=400)
        
    except Exception as e:
        st.error(f"錯誤：{e}")

if not uploaded_file:
    st.info("請上傳 PDF")
