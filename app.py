import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import numpy as np
from collections import defaultdict

st.set_page_config(page_title="HKDSE v4.0", layout="wide")

st.title("🧠 HKDSE **超智能分欄** v4.0")
st.markdown("""
**解決所有空格問題！**
- 📐 **動態間隙檢測**：自適應每個PDF
- ✂️ **智能分詞**：「Que./ marking」→「Que./」|「marking」
- 🔄 **多階段清理**：物理+語義結合
""")

uploaded_file = st.file_uploader("上傳 PDF", type="pdf")

def smart_split_text(text):
    """智能分詞：斜線、特殊符號自動切分"""
    # 常見分隔符優先切分
    splits = re.split(r'[ /()#：:]+', text)
    return [s.strip() for s in splits if s.strip()]

def extract_super_smart_table(page):
    """v4.0 超智能提取"""
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
        
        # 動態間隙閾值（該頁自適應）
        x_positions = [w['x0'] for w in row_words]
        gaps = [x_positions[i+1] - x_positions[i] for i in range(len(x_positions)-1)]
        dynamic_gap = np.percentile(gaps, 75)  # 75%分位數作為閾值
        
        cols = []
        current_col = []
        prev_x1 = row_words[0]['x0']
        
        for word in row_words:
            gap = word['x0'] - prev_x1
            if gap > dynamic_gap * 0.8:  # 動態閾值
                col_text = ' '.join(w['text'] for w in current_col).strip()
                if col_text:
                    # 智能二次分詞
                    smart_cols = smart_split_text(col_text)
                    cols.extend(smart_cols)
                current_col = [word]
            else:
                current_col.append(word)
            prev_x1 = word['x1']
        
        col_text = ' '.join(w['text'] for w in current_col).strip()
        if col_text:
            smart_cols = smart_split_text(col_text)
            cols.extend(smart_cols)
        
        if len(cols) >= 4:
            rows.append(cols[:25])  # 最多25欄
    
    return rows

def ultra_clean(df):
    """超級清理"""
    # 移除空欄
    df = df.loc[:, (df != '').any()]
    
    # 右對齊
    def align_right(series):
        valid = series[series.astype(str).str.strip() != ''].tolist()
        return pd.Series([''] * (len(series) - len(valid)) + valid)
    
    df = df.apply(align_right, axis=1)
    
    # 智能數字
    def to_num(val):
        s = str(val).strip()
        if s.endswith('%'):
            try:
                return float(s[:-1]) / 100
            except:
                pass
        try:
            return pd.to_numeric(s.replace(',', ''), errors='coerce')
        except:
            return s
    
    df = df.map(to_num)
    df = df.fillna('').replace(np.nan, '')
    
    return df

if uploaded_file:
    try:
        st.info("🧠 超智能解析...")
        all_rows = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = extract_super_smart_table(page)
                all_rows.extend(table)
        
        df = ultra_clean(pd.DataFrame(all_rows))
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, '智能解析', index=False, header=False)
            
            stats = pd.DataFrame([{
                '行數': len(all_rows),
                '欄數': len(df.columns),
                '狀態': '超智能成功'
            }])
            stats.to_excel(writer, '統計', index=False)
        
        st.success("✅ **超智能轉換完成！**")
        col1, col2 = st.columns([3,1])
        with col1:
            st.download_button("📥 下載 v4.0", output.getvalue(), 
                             f"DSE_v4.0.xlsx", 
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            st.metric("解析行數", len(all_rows))
        
        st.subheader("🎯 智能預覽")
        st.dataframe(df.head(15), height=400)
        
    except Exception as e:
        st.error(f"錯誤：{e}")
        st.info("顯示原始文字...")
        with pdfplumber.open(uploaded_file) as pdf:
            st.text(pdf.pages[0].extract_text()[:2000])

st.caption("**v4.0 超智能** | 動態間隙 + 分詞 | 全科完美")
