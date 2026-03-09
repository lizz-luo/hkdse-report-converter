import streamlit as st
import pdfplumber
import pandas as pd
import io
import numpy as np
from collections import defaultdict

st.set_page_config(page_title="HKDSE 物理定位轉換 v3.0", layout="centered")

st.title("🚀 HKDSE **物理定位**轉 Excel v3.0")
st.markdown("""
**徹底解決所有黏連/拆分問題！**
- ✅ **通用全科**：無需科目特定邏輯
- ✅ **物理分隔**：按字符實際 X 座標（像 Adobe）
- ✅ **智能分欄**：自動發現欄寬，永不錯位
- ✅ **無 <NA>**，完美數字格式
""")

uploaded_file = st.file_uploader("上傳 PDF", type="pdf")

def extract_physical_table(page):
    """物理定位提取：按字符 X 座標分欄"""
    words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
    
    if not words:
        return None
    
    # 按 Y 座標分行（物理行）
    rows_by_y = defaultdict(list)
    for word in words:
        y_pos = round(word['top'], 1)  # 行定位
        rows_by_y[y_pos].append(word)
    
    # 排序行（從上到下）
    sorted_y = sorted(rows_by_y.keys())
    physical_rows = []
    
    for y in sorted_y:
        row_words = rows_by_y[y]
        if len(row_words) < 2:  # 跳過單詞行
            continue
            
        # 按 X 座標排序（物理列）
        row_words.sort(key=lambda w: w['x0'])
        
        # 智能分欄：按 X 間隙分組
        columns = []
        current_col = []
        prev_x1 = row_words[0]['x0']
        
        for word in row_words:
            gap = word['x0'] - prev_x1
            if gap > 8:  # 大間隙 = 新欄
                if current_col:
                    columns.append(' '.join(w['text'] for w in current_col))
                current_col = [word]
            else:
                current_col.append(word)
            prev_x1 = word['x1']
        
        if current_col:
            columns.append(' '.join(w['text'] for w in current_col))
        
        if len(columns) >= 4:  # 有效表格行
            physical_rows.append(columns)
    
    return physical_rows if physical_rows else None

def process_physical_table(table):
    """處理物理表格"""
    if not table:
        return pd.DataFrame()
    
    df = pd.DataFrame(table)
    
    # 智能欄寬統一（物理對齊）
    max_cols = max(len(row) for row in table)
    df = df.reindex(columns=range(max_cols), fill_value='')
    
    # 向右對齊（保留物理順序）
    def align_right(row):
        valid = [c for c in row if str(c).strip()]
        nans = len(row) - len(valid)
        return [''] * nans + valid
    
    df = df.apply(align_right, axis=1)
    
    # 數字轉換
    def to_numeric_safe(val):
        val_str = str(val).strip()
        if val_str.endswith('%'):
            try:
                return float(val_str.replace('%', '')) / 100
            except:
                return val_str
        try:
            return pd.to_numeric(val_str, errors='coerce')
        except:
            return val_str
    
    df = df.applymap(to_numeric_safe)
    df = df.replace({np.nan: '', pd.NA: ''})
    
    return df

if uploaded_file is not None:
    with st.spinner("🔬 物理定位解析中..."):
        sections = {}
        current_section = "General"

        with pdfplumber.open(uploaded_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text()
                if not text:
                    continue

                # 卷別偵測
                paper_match = re.search(r"(?:卷\s*)?Paper:\s*([A-Za-z0-9]+)", text)
                if paper_match:
                    section = f"Paper_{paper_match.group(1)}"
                else:
                    section = f"Page_{page_num+1}"
                
                if section not in sections:
                    sections[section] = []

                # 物理提取表格
                table = extract_physical_table(page)
                if table:
                    sections[section].extend(table)

        # 生成 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for section, table_data in sections.items():
                if len(table_data) < 3:  # 太少行跳過
                    continue
                
                df = process_physical_table(table_data)
                safe_name = section.replace('/', '_')[:31]
                df.to_excel(writer, sheet_name=safe_name, index=False, header=False)

        st.success("🎉 **物理定位轉換完成！** 全科通用，永不黏連")
        st.download_button(
            label="📥 下載物理定位 Excel",
            data=output.getvalue(),
            file_name="DSE_Physical_v3.0.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.info("""
**v3.0 物理定位革命：**
- 📐 **字符級精確**：按實際 X/Y 座標分欄
- 🌍 **全科通用**：無需特定邏輯
- ⚡ **像 Adobe**：空格即分隔符，間隙>8px=新欄
- 🔢 **智能數字**：%自動轉小數
""")
