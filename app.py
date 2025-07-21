import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# === Page Config ===
st.set_page_config(page_title="PDF to Excel Converter", layout="centered")

# === Title and Description ===
st.title("📄 AIIMS PDF to Excel Converter")
st.markdown("Upload a multi-page PDF containing tables. This tool extracts all tables and converts them into an Excel file for download.")

# === File Upload ===
uploaded_pdf = st.file_uploader("📎 Upload PDF File", type=["pdf"])

# === Function to Make Headers Unique ===
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

# === Main Processing ===
if uploaded_pdf:
    with st.spinner("⏳ Extracting tables from PDF... Please wait."):
        all_tables = []
        try:
            with pdfplumber.open(uploaded_pdf) as pdf:
                total_pages = len(pdf.pages)
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) > 1:
                            headers = table[0]
                            unique_headers = make_columns_unique(headers)
                            df = pd.DataFrame(table[1:], columns=unique_headers)
                            all_tables.append(df)

            if all_tables:
                final_df = pd.concat(all_tables, ignore_index=True)

                # Save to Excel in memory
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Extracted Data')
                    writer.save()
                    processed_data = output.getvalue()

                st.success("✅ Tables extracted and Excel file ready!")
                st.download_button(
                    label="📥 Download Excel File",
                    data=processed_data,
                    file_name="AIIMS_Paramedical_Result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ No tables were found in the uploaded PDF.")
        except Exception as e:
            st.error(f"❌ An error occurred: {e}")
