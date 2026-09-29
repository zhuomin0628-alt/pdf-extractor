import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="PDF 零件長度擷取器", layout="centered")

st.title("📄 PDF 零件編號與最長尺寸擷取工具")
st.write(
    "支援**單張**或**合併 PDF** 檔案。系統將自動逐頁掃描，抓取圖面內部的**零件編號 (MARK)** 與**最長尺寸線**。"
)

# 檔案上傳元件（支援多選）
uploaded_files = st.file_uploader(
    "選擇 PDF 檔案 (可單選或多選)", type=["pdf"], accept_multiple_files=True
)


def extract_pdf_data(file_obj):
  results = []
  # 使用 pdfplumber 逐頁開啟並解析 PDF
  with pdfplumber.open(file_obj) as pdf:
    for page_num, page in enumerate(pdf.pages, start=1):
      text = page.extract_text()
      if not text:
        continue

      # 1. 抓取圖面內部出現的零件編號 (例如 5CP16, 1CP91, 6CP1 等)
      mark_matches = re.findall(
          r"([A-Z0-9]+CP[0-9]+)", text, re.IGNORECASE
      )

      if mark_matches:
        # 取該頁面找到的第一個（或主要的）零件編號作為代表
        mark_name = mark_matches[0].upper().replace(".", "")
      else:
        # 如果該頁剛好沒抓到標記，則以頁碼命名避免遺漏
        mark_name = f"PAGE_{page_num}"

      # 2. 抓取圖面上的所有尺寸數值 (尋找 3 到 4 位數、可帶一位小數的數字)
      number_matches = re.findall(r"\b([0-9]{3,4}\.[0-9]|[0-9]{4})\b", text)

      valid_lengths = []
      for num_str in number_matches:
        val = float(num_str)
        # 嚴格篩選長度範圍：限制在 800 到 4000 之間，並排除年份 2026 避免干擾
        if 800 <= val <= 4000 and val != 2026:
          valid_lengths.append(val)

      # 3. 取得該頁面中最長的尺寸數值作為圖面總長度
      max_length = max(valid_lengths) if valid_lengths else 0.0

      if max_length > 0:
        results.append({
            "零件編號 (MARK)": mark_name,
            "圖面最長尺寸": max_length,
            "來源頁面": f"第 {page_num} 頁",
        })

  return results


if uploaded_files:
  if st.button("🚀 開始解析 PDF 資料"):
    all_data = []

    with st.spinner("正在逐頁解析 PDF 內容與尺寸，請稍候..."):
      for uploaded_file in uploaded_files:
        file_results = extract_pdf_data(uploaded_file)
        all_data.extend(file_results)

    if all_data:
      df = pd.DataFrame(all_data)

      # 顯示解析結果預覽
      st.success(f"解析成功！總共抓取到 {len(df)} 筆零件尺寸資料。")
      st.dataframe(df)

      # 轉換成 Excel 供下載
      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="PDF資料")
      processed_data = output.getvalue()

      st.download_button(
          label="📥 下載 Excel 對照表",
          data=processed_data,
          file_name="PDF零件長度擷取結果.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )
    else:
      st.warning(
          "未能從上傳的 PDF 中抓取到符合條件的零件與尺寸，請確認 PDF"
          " 是否為含有文字圖層的工程圖。"
      )
