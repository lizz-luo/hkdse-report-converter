import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import numpy as np
from collections import defaultdict

st.set_page_config(page_title="HKDSE v4.1", layout="wide")

st.title("🎯 HKDSE **完整智能版** v4.1")
st.markdown("""
**智能分詞 + 右對齊 + 刪空欄 全自動**
- ✂️ 動態分詞（Que./ marking → Que./ | marking）
- ➡️ **完美右對齊**
- 🗑️ **自動刪左空欄**
""")

uploaded_file = st.file_uploader("上傳 PDF", type="pdf")

def smart_split_text(text):
    """智能分詞"""
    splits = re.split(r'[ /()#：:]+', text)
    return [s.strip() for s in splits if s.strip()]

def extract_smart_table(page):
    """超智能提取"""
    words = page.extract_words(x_tolerance=1.5, y_tolerance=2.5)
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
        
        # 動態間隙
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

def perfect_alignment(df):
    """**完美對齊流程**"""
    # 1. 刪除完全空欄
    df = df.loc[:, (df != '').any(axis=0)]
    
    # 2. 右對齊（保留有效數據）
    def right_align(series):
        valid = series[series.astype(str).str.strip() != ''].tolist()
        n_empty = len(series) - len(valid)
        return pd.Series([''] * n_empty + valid)
    
    df = df.apply(right_align, axis=1)
    
    # 3. 再次刪除**新增的左空欄**
    df = df.loc[:, (df != '').any(axis=0)]
    
    return df

def smart_numeric(val):
    s = str(val).strip()
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
        st.info("🎯 智能解析 + 完美對齊...")
        all_rows = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = extract_smart_table(page)
                all_rows.extend(table)
        
        # **核心三步：分詞 → 右對齊 → 刪空欄**
        df_raw = pd.DataFrame(all_rows)
        df_aligned = perfect_alignment(df_raw)
        df_final = df_aligned.map(smart_numeric).fillna('')
        
        # 輸出
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, '完美對齊', index=False, header=False)
            
            stats = pd.DataFrame([{
                '原始行': len(all_rows),
                '最終欄': len(df_final.columns),
                '有效數據': (df_final != '').sum().sum(),
                '狀態': '完美對齊完成'
            }])
            stats.to_excel(writer, '統計', index=False)
        
        st.success("✅ **完美對齊完成！**")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button("📥 下載 v4.1", output.getvalue(), 
                             "DSE_v4.1_Perfect.xlsx")
        with col2:
            st.metric("解析行數", len(all_rows))
        with col3:
            st.metric("最終欄數", len(df_final.columns))
        
        # **對比展示**
        st.subheader("📊 對齊前後對比")
        col_a, col_b = st.columns(2)
        with col_a:
            st.write("**右對齊前**")
            st.dataframe(df_raw.head(8), height=300)
        with col_b:
            st.write("**完美對齊後**")
            st.dataframe(df_final.head(8), height=300)
            
    except Exception as e:
        st.error(f"錯誤：{e}")

st.caption("**v4.1 | 智能分詞 + 完美對齊流程**")
