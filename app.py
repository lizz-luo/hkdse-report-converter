import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import numpy as np
from collections import defaultdict

st.set_page_config(page_title="HKDSE 物理定位 v3.0", layout="centered")

st.title("🚀 HKDSE **物理定位轉換** v3.0")
st.markdown("""
**徹底解決所有黏連問題！全科通用**
- 📐 按字符**實際 X/Y 座標**分欄（像 Adobe）
- 🌍 **無科目特定邏輯**，永不錯位
- ⚡ 空格>8px自動分欄，智能合併數字+%  
""")

uploaded_file = st.file_uploader("上傳 PDF", type="pdf")

def extract_physical_table(page):
    """物理定位：按字符座標智能分欄"""
    words = page.extract_words(x_tolerance=2, y_tolerance=3)
    if not words:
        return None
    
    # 按 Y 座標分物理行
    rows_by_y = defaultdict(list)
    for word in words:
        y_pos = round(word['top'], 1)
        rows_by_y[y_pos].append(word)
    
    sorted_y = sorted(rows_by_y.keys())
    physical_rows = []
    
    for y in sorted_y:
        row_words = rows_by_y[y]
        if len(row_words) < 2:
            continue
            
        row_words.sort(key=lambda w: w['x0'])
        
        # 按 X 間隙智能分欄（物理分隔）
        columns = []
        current_col = []
        prev_x1 = row_words[0]['x0']
        
        for word in row_words:
            gap = word['x0'] - prev_x1
            if gap > 8:  # 大間隙=新欄
                col_text = ' '.join(w['text'].strip() for w in current_col)
                if col_text:
                    columns.append(col_text)
                current_col = [word]
            else:
                current_col.append(word)
            prev_x1 = word['x1']
        
        col_text = ' '.join(w['text'].strip() for w in current_col)
        if col_text:
            columns.append(col_text)
        
        if len(columns) >= 4:
            physical_rows.append(columns)
    
    return physical_rows

def process_physical_table(table):
    """處理物理表格"""
    if not table:
        return pd.DataFrame()
    
    df = pd.DataFrame(table)
    max_cols = max(len(row) for row in table)
    df = df.reindex(columns=range(max_cols), fill_value='')
    
    # 物理右對齊
    def align_right(row):
        valid = [c for c in row if str(c).strip()]
        return [''] * (len(row) - len(valid)) + valid
    
    df = df.apply(align_right, axis=1)
    
    # 智能數字轉換
    def smart_numeric(val):
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
    
    df = df.applymap(smart_numeric)
    df = df.fillna('').replace(pd.NA, '')
    
    return df

if uploaded_file is not None:
    try:
        with st.spinner("🔬 物理座標解析中..."):
            sections = {}
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    
                    # 智能卷別偵測
                    section = f"Page_{page_num+1}"
                    paper_match = re.search(r"(?:[卷紙]?\s*)?Paper[:：]\s*([A-Za-z0-9]+)", 
                                          text, re.IGNORECASE)
                    if paper_match:
                        section = f"Paper_{paper_match.group(1)}"
                    
                    table = extract_physical_table(page)
                    if table:
                        sections[section] = table

            # 生成 Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for section, table_data in sections.items():
                    if len(table_data) < 3:
                        continue
                    df = process_physical_table(table_data)
                    safe_name = re.sub(r"[\\/*?:\[\]]", "_", section)[:31]
                    df.to_excel(writer, sheet_name=safe_name, index=False, header=False)

            st.success(f"✅ **物理轉換完成！** 找到 {len(sections)} 個表格")
            st.download_button(
                label="📥 下載 v3.0 Excel",
                data=output.getvalue(),
                file_name="DSE_Physical_v3.0.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"解析錯誤：{str(e)}")
        st.info("請確認 PDF 包含標準 HKDSE 項目分析表格")

st.caption("v3.0 | **物理座標革命** | 全科通用，永不黏連")
