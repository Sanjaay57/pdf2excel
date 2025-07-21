import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# === Page Configuration ===
st.set_page_config(page_title="PDF to Excel Converter", layout="centered")

# === Title ===
st.title("📄PDF to Excel Converter")
st.markdown("""
Upload your AIIMS Paramedical result PDF (multi-page supported).
The tables will be extracted and converted into an Excel file, named automatically based on your uploaded PDF.
""")

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

# === Main Logic ===
if uploaded_pdf:
    with st.spinner("⏳ Extracting tables from PDF..."):
        all_tables = []
        try:
            with pdfplumber.open(uploaded_pdf) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) > 1:
                            headers = make_columns_unique(table[0])
                            df = pd.DataFrame(table[1:], columns=headers)
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
                st.warning("⚠️ No tables were found in this PDF.")

        except Exception as e:
            st.error(f"❌ An error occurred: {e}")
