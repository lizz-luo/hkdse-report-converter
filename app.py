import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import numpy as np

st.set_page_config(page_title="HKDSE 全科數據轉換工具", layout="centered")

st.title("📊 HKDSE 項目分析報告轉 Excel 工具")
st.write(
    "請上傳考評局的 PDF 報告（支援中英數等各科）。\n"
    "系統會自動：\n"
    "1. 依座標提取表格並按卷別分 Sheet\n"
    "2. 智能修補「106 被拆成 10 和 6」等錯位\n"
    "3. 將數據向右對齊、刪除多餘空欄\n"
    "4. 將能轉換的數值變成 Excel 中可計算的數字格式\n"
    "5. 空位顯示為空白（非 <NA>）"
)

uploaded_file = st.file_uploader("請點擊或拖曳上傳 PDF 檔案", type="pdf")


def clean_and_convert_to_numeric(val):
    """將文字轉換為真實數字的智能函數"""
    if pd.isna(val):
        return pd.NA

    val_str = str(val).strip()
    val_str = val_str.replace(',', '')

    if val_str.endswith('%'):
        try:
            return float(val_str.replace('%', '')) / 100.0
        except ValueError:
            return val_str

    if val_str.startswith('+'):
        val_str = val_str[1:]

    try:
        num = float(val_str)
        return int(num) if num.is_integer() else num
    except ValueError:
        return val_str


def fix_row_split_numbers(cells):
    """
    行級修補邏輯：自動合併被誤切的數字（例如 106 → 10 + 6）
    """
    cells = [("" if c is None else str(c)) for c in cells]

    for i in range(len(cells) - 2):
        a, b, c = cells[i], cells[i + 1], cells[i + 2]

        if not (a and b and c):
            continue

        # a、b 都是純數字，a 兩位數 b 一位數，且 c 像是 100%
        if a.isdigit() and b.isdigit() and len(a) == 2 and len(b) == 1:
            c_clean = c.replace(',', '')
            if c_clean.startswith("100"):
                merged = a + b  # "10" + "6" → "106"
                cells[i] = merged
                cells[i + 1] = ""  # 留空，後續向右對齊會處理
                break

    return cells


if uploaded_file is not None:
    with st.spinner("正在解析 PDF 並處理數據，請稍候..."):
        sections = {}
        current_section = "General"
        detected_subject = "General"

        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                # 科目偵測
                if ("Chinese Language" in text) or ("中國語文" in text):
                    detected_subject = "Chinese"
                elif ("English Language" in text) or ("英國語文" in text):
                    detected_subject = "English"
                elif ("Mathematics" in text) or ("數學" in text):
                    detected_subject = "Math"

                # 卷別偵測
                paper_match = re.search(r"(?:卷\s*)?Paper:\s*([A-Za-z0-9]+)", text)
                if paper_match:
                    paper_name = paper_match.group(1).strip()
                    current_section = f"Paper_{paper_name}"
                else:
                    if current_section == "General" and detected_subject != "General":
                        current_section = f"{detected_subject}_General"

                if current_section not in sections:
                    sections[current_section] = []

                # 動態調整 X 軸容錯率
                x_tolerance = 2 if detected_subject == "Chinese" else 3

                dynamic_table_settings = {
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "intersection_x_tolerance": x_tolerance,
                    "intersection_y_tolerance": 2,
                    "min_words_vertical": 2,
                }

                table = page.extract_table(dynamic_table_settings)

                if table:
                    sections[current_section].extend(table)
                else:
                    fallback_table = page.extract_table()
                    if fallback_table:
                        sections[current_section].extend(fallback_table)

        # 整批清洗與處理
        output = io.BytesIO()
        has_data = False

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for section_name, data in sections.items():
                if not data:
                    continue

                # 去掉完全空白的行
                cleaned_rows = []
                for row in data:
                    if any((isinstance(cell, str) and cell.strip()) for cell in row):
                        cleaned_rows.append(row)

                if not cleaned_rows:
                    continue

                df = pd.DataFrame(cleaned_rows)

                # 將空字串 / 空白轉為 NA
                df.replace(r"^\s*$", pd.NA, regex=True, inplace=True)
                df.dropna(how='all', axis=1, inplace=True)

                # 行級數位修補
                df = df.apply(lambda row: fix_row_split_numbers(list(row)), axis=1, result_type="expand")

                # 再次統一空值
                df.replace(r"^\s*$", pd.NA, regex=True, inplace=True)

                # 向右對齊
                def shift_row_right(row):
                    valid_vals = [v for v in row if pd.notna(v) and str(v).strip() != ""]
                    num_nans = len(row) - len(valid_vals)
                    return pd.Series([pd.NA] * num_nans + valid_vals, index=row.index)

                df = df.apply(shift_row_right, axis=1)

                # 刪除因對齊產生的左側全空欄
                df.dropna(how='all', axis=1, inplace=True)

                # 數字轉換
                df = df.applymap(clean_and_convert_to_numeric)

                # 【最終修正】：將 NA 替換為空字串，避免 Excel 顯示 <NA>
                df = df.fillna('')

                safe_sheet_name = re.sub(r"[\\/*?:\[\]]", "_", section_name)[:31]
                df.to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)
                has_data = True

        if has_data:
            st.success(
                f"✅ 轉換完成！已偵測科目：{detected_subject}，並套用混合策略修補與數字格式。\n"
                "空位已顯示為空白格（非 <NA>）。"
            )
            st.download_button(
                label="📥 下載 Excel 檔案",
                data=output.getvalue(),
                file_name="DSE_Report_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error("⚠️ 未能從上傳的 PDF 中提取到有效表格數據。")
