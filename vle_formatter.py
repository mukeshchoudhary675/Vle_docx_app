import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from io import BytesIO

st.title("📄 VLE Data Formatter")

uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("Excel uploaded successfully!")
    st.dataframe(df.head())

    columns = st.multiselect("🧩 Select columns to include", df.columns.tolist(), default=df.columns.tolist()[:3])
    font_size = st.slider("🔠 Font size", min_value=10, max_value=24, value=14)
    records_per_page = st.slider("📃 Records per page", min_value=1, max_value=5, value=1)

    if st.button("🛠️ Generate DOCX"):
        doc = Document()
        for i in range(0, len(df), records_per_page):
            chunk = df.iloc[i:i+records_per_page]
            for _, row in chunk.iterrows():
                # Column name mapping (Excel column → Word label)
                column_mapping = {
                    "Name of VLE": "Name",
                    "VLE  Contact No.": "Contact No.",
                    "VLE Address": "Address"
                }
                
                for col in columns:
                    display_name = column_mapping.get(col, col)  # Use mapped name if available, else original
                    p = doc.add_paragraph()
                    run = p.add_run(f"{display_name}: {row[col]}")
                    run.font.size = Pt(font_size)

                # for col in columns:
                #     p = doc.add_paragraph()
                #     run = p.add_run(f"{col.upper()}: {row[col]}")
                #     run.font.size = Pt(font_size)
                doc.add_paragraph()
            doc.add_page_break()

        output = BytesIO()
        doc.save(output)
        output.seek(0)

        st.success("✅ Document generated!")
        st.download_button("📥 Download DOCX", data=output, file_name="VLE_Output.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
