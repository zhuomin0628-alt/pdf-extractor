import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="鋼構圖面尺寸自動擷取工具", page_icon="📐", layout="centered"
)

st.title("📐 鋼構 PDF 零件編號與最長尺寸自動擷取")
st.write(
    "請上傳您的 PDF 零件圖或清單檔案，系統將自動解析並計算每個零件編號的最長尺寸。"
)

uploaded_files = st.file_uploader(
    "選擇 PDF 檔案", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
  all_records = []

  with st.spinner("正在解析 PDF 檔案，請稍候..."):
    for uploaded_file in uploaded_files:
      file_name = uploaded_file.name
      default_mark = file_name.split(".")[0].replace("_", "").upper()

      pdf_bytes = uploaded_file.read()
      max_len = 0

      with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full_text = ""
        for page in pdf.pages:
          text = page.extract_text()
          if text:
            full_text += text + "\n"

        matches = re.findall(r"([0-9]{3,4}\.[0-9]|[0-9]{4})", full_text)
        for m in matches:
          val = float(m)
          if 500 <= val <= 4000 and val != 2026:
            if val > max_len:
              max_len = val

      if max_len > 0:
        all_records.append(
            {"零件編號 (MARK)": default_mark, "最大尺寸 (LENGTH)": max_len}
        )

  if all_records:
    df = pd.DataFrame(all_records)
    df = df.groupby("零件編號 (MARK)", as_index=False)["最大尺寸 (LENGTH)"].max()

    st.success(f"成功解析！總共取得 {len(df)} 筆零件資料。")
    st.dataframe(df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df.to_excel(writer, index=False, sheet_name="PDF尺寸")
    excel_data = output.getvalue()

    st.download_button(
        label="📥 下載整理好的 Excel 尺寸表",
        data=excel_data,
        file_name="PDF_Extracted_Lengths.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
  else:
    st.warning("未能從上傳的 PDF 中偵測到有效的尺寸數據，請檢查檔案格式。")
