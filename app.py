import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import numpy as np

st.set_page_config(page_title="HKDSE 全科數據轉換工具", layout="centered")

st.title("📊 HKDSE 項目分析報告轉 Excel 工具")
st.write(
    "請上傳考評局的 PDF 報告（支援多科目）。\n"
    "系統會自動：\n"
    "1. 依座標提取表格並按卷別分 Sheet\n"
    "2. 智能修補「106 被拆成 10 和 6」等錯位\n"
    "3. 將數據向右對齊、刪除多餘空欄\n"
    "4. 將能轉換的數值變成 Excel 中可計算的數字格式"
)

uploaded_file = st.file_uploader("請點擊或拖曳上傳 PDF 檔案", type="pdf")


def clean_and_convert_to_numeric(val):
    """將文字轉換為真實數字的智能函數。"""
    if pd.isna(val):
        return val

    val_str = str(val).strip()
    # 去掉千分位逗號
    val_str = val_str.replace(",", "")

    # 處理百分比：例如 "85%" -> 0.85
    if val_str.endswith("%"):
        try:
            return float(val_str.replace("%", "")) / 100.0
        except ValueError:
            return val_str

    # 去掉開頭的 "+" 號（如 +0.25 -> 0.25）
    if val_str.startswith("+"):
        val_str = val_str[1:]

    # 嘗試轉成數字
    try:
        num = float(val_str)
        return int(num) if num.is_integer() else num
    except ValueError:
        return val_str


def fix_row_split_numbers(cells):
    """
    行級修補邏輯：
    針對像「106 被切成 '10'、'6'，下一欄是 100 或 100.0」這類情況，
    自動把 '10' 和 '6' 合併為 '106'，把原本第二個位置清空。
    """
    # 轉成純字串方便檢查
    cells = [("" if c is None else str(c)) for c in cells]

    for i in range(len(cells) - 2):
        a, b, c = cells[i], cells[i + 1], cells[i + 2]

        # 三個都非空才有可能
        if not (a and b and c):
            continue

        # a、b 都是純數字，且 a 為兩位數，b 為一位數
        if a.isdigit() and b.isdigit() and len(a) == 2 and len(b) == 1:
            # c 看起來像是 100% 或 100.0 這類百分比數
            c_clean = c.replace(",", "")
            if c_clean.startswith("100"):
                merged = a + b  # "10" + "6" -> "106"
                cells[i] = merged
                cells[i + 1] = ""  # 留空，之後的「向右對齊 + 刪空欄」會處理掉
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

                # 科目偵測，用於決定切欄的容錯策略
                if ("Chinese Language" in text) or ("中國語文" in text):
                    detected_subject = "Chinese"
                elif ("English Language" in text) or ("英國語文" in text):
                    detected_subject = "English"
                elif ("Mathematics" in text) or ("數學" in text):
                    detected_subject = "Math"

                # 卷別偵測（例如 Paper: 1A, Paper: 2 ...)
                paper_match = re.search(r"(?:卷\s*)?Paper:\s*([A-Za-z0-9]+)", text)
                if paper_match:
                    paper_name = paper_match.group(1).strip()
                    current_section = f"Paper_{paper_name}"
                else:
                    if current_section == "General" and detected_subject != "General":
                        current_section = f"{detected_subject}_General"

                if current_section not in sections:
                    sections[current_section] = []

                # 依科目動態調整 X 軸容錯率
                # 中文科排版較密，需嚴謹（2）；英文、數學略鬆（3）
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
                    # 備用：用預設參數再嘗試一次
                    fallback_table = page.extract_table()
                    if fallback_table:
                        sections[current_section].extend(fallback_table)

        # 第二階段：整批清洗、右對齊、修補與數字轉換，並輸出 Excel
        output = io.BytesIO()
        has_data = False

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
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
                # 先刪掉全空欄（這一步可以去掉很多幽靈欄）
                df.dropna(how="all", axis=1, inplace=True)

                # 對每一行先套用「拆錯數字修補」邏輯
                df = df.apply(lambda row: fix_row_split_numbers(list(row)), axis=1, result_type="expand")

                # 再次統一空值
                df.replace(r"^\s*$", pd.NA, regex=True, inplace=True)

                # 向右對齊：把 NA 推到左邊，真實數值推到右邊
                def shift_row_right(row):
                    valid_vals = [v for v in row if pd.notna(v) and str(v).strip() != ""]
                    num_nans = len(row) - len(valid_vals)
                    return pd.Series([pd.NA] * num_nans + valid_vals, index=row.index)

                df = df.apply(shift_row_right, axis=1)

                # 再刪一次「全空欄」，特別是最左邊那幾欄
                df.dropna(how="all", axis=1, inplace=True)

                # 將每個儲存格嘗試轉成數字
                df = df.applymap(clean_and_convert_to_numeric)

                safe_sheet_name = re.sub(r"[\\/*?:\[\]]", "_", section_name)[:31]
                df.to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)
                has_data = True

        if has_data:
            st.success(
                f"✅ 轉換完成！已偵測科目：{detected_subject}，並套用混合策略修補與數字格式。\n"
                "請下載下方的 Excel 檔案，並歡迎用不同 PDF 測試、再把遇到的特例告訴我，我們可以繼續微調規則。"
            )
            st.download_button(
                label="📥 下載 Excel 檔案",
                data=output.getvalue(),
                file_name="DSE_Report_Mixed_Strategy.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error("⚠️ 未能從上傳的 PDF 中提取到有效表格數據。")
