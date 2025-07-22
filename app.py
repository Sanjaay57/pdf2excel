import streamlit as st
import pdfplumber
import pandas as pd
import pytesseract
from PIL import Image
from pdf2image import convert_from_bytes
from io import BytesIO
import tempfile

# === Page Configuration ===
st.set_page_config(page_title="PDF to Excel Converter with OCR", layout="centered")
st.title("📄 PDF to Excel Converter with OCR")
st.markdown("""
Upload your AIIMS Paramedical result PDF (multi-page supported).<br>
It will extract tables, and if the page is scanned, OCR will be used automatically.
""", unsafe_allow_html=True)

# === File Upload ===
uploaded_pdf = st.file_uploader("📎 Upload PDF File", type=["pdf"])

# === Helper to Ensure Unique Headers ===
def make_columns_unique(columns):
    seen = {}
    new_columns = []
    for col in columns:
        if col is None or col.strip() == "":
            col = "Unnamed"
        if col in seen:
            seen[col] += 1
            new_columns.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            new_columns.append(col)
    return new_columns

# === OCR Table Extraction Helper ===
def extract_ocr_table_from_image(img: Image.Image) -> pd.DataFrame:
    ocr_data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DATAFRAME)
    ocr_data = ocr_data.dropna(subset=['text'])
    if ocr_data.empty:
        return pd.DataFrame()
    
    lines = ocr_data.groupby('line_num')['text'].apply(lambda x: ' '.join(x)).tolist()
    table_lines = [line for line in lines if any(char.isdigit() for char in line)]

    if not table_lines or len(table_lines) < 2:
        return pd.DataFrame()

    headers = table_lines[0].split()
    rows = [line.split() for line in table_lines[1:] if len(line.split()) == len(headers)]
    df = pd.DataFrame(rows, columns=make_columns_unique(headers))
    return df

# === Main Logic ===
if uploaded_pdf:
    with st.spinner("⏳ Extracting tables (including OCR fallback)..."):
        all_tables = []

        try:
            text_based_pages = set()
            with pdfplumber.open(uploaded_pdf) as pdf:
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            if table and len(table) > 1:
                                headers = make_columns_unique(table[0])
                                df = pd.DataFrame(table[1:], columns=headers)
                                all_tables.append(df)
                        text_based_pages.add(i)

            # Identify image-only pages and apply OCR
            pdf_bytes = uploaded_pdf.read()
            images = convert_from_bytes(pdf_bytes, dpi=300)
            for i, img in enumerate(images):
                if i not in text_based_pages:
                    df = extract_ocr_table_from_image(img)
                    if not df.empty:
                        all_tables.append(df)

            if all_tables:
                final_df = pd.concat(all_tables, ignore_index=True)

                # Save to Excel in memory
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Extracted Data')
                output.seek(0)
                processed_data = output.getvalue()

                # Dynamically name Excel file after uploaded PDF
                pdf_name = uploaded_pdf.name.rsplit(".", 1)[0]
                excel_name = f"{pdf_name}.xlsx"

                st.success(f"✅ Extraction complete! Click below to download **{excel_name}**")
                st.download_button(
                    label="📥 Download Excel File",
                    data=processed_data,
                    file_name=excel_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ No tables or OCR results were found in this PDF.")

        except Exception as e:
            st.error(f"❌ An error occurred: {e}")
