# app.py
import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="PDF to Excel Converter", layout="centered")

st.title("📄PDF to Excel Converter")
st.write("Upload a multi-page PDF with tables. This app will extract the tables and export them to Excel.")

uploaded_pdf = st.file_uploader("Upload your PDF file", type=["pdf"])

if uploaded_pdf:
    with st.spinner("Processing PDF..."):
        all_tables = []
        try:
            with pdfplumber.open(uploaded_pdf) as pdf:
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for table in tables:
                        if table:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            all_tables.append(df)
            if all_tables:
                final_df = pd.concat(all_tables, ignore_index=True)

                # Save to Excel in memory
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Extracted Data')
                    writer.save()
                    processed_data = output.getvalue()

                st.success("✅ Tables extracted successfully!")
                st.download_button(
                    label="📥 Download Excel",
                    data=processed_data,
                    file_name="AIIMS_Extracted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No tables found in the uploaded PDF.")

        except Exception as e:
            st.error(f"An error occurred: {e}")
