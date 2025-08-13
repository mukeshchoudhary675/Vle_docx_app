import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from io import BytesIO

st.title("📄 VLE Data Formatter (Dynamic Settings & Live Preview)")

# Upload Excel
uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ Excel uploaded successfully!")
    st.dataframe(df.head())

    # Step 1: Select columns to include
    selected_columns = st.multiselect(
        "🧩 Select columns to include in the output",
        df.columns.tolist(),
        default=df.columns.tolist()[:3]
    )

    # Step 2: Dynamic renaming inputs
    st.subheader("✏ Rename Columns for Output")
    column_rename_map = {}
    for col in selected_columns:
        new_name = st.text_input(f"Rename '{col}' to:", value=col)
        column_rename_map[col] = new_name

    # Step 3: Choose bold fields
    bold_fields = st.multiselect(
        "🖋 Select fields to make BOLD in output",
        list(column_rename_map.values())
    )

    # Step 4: Font size and layout settings
    font_size = st.slider("🔠 Font size", min_value=10, max_value=30, value=14)
    records_per_page = st.slider("📃 Records per page", min_value=1, max_value=5, value=1)

    # Step 5: Text case option
    case_option = st.selectbox(
        "🔡 Text casing",
        ["Original", "UPPERCASE", "Proper Case", "lowercase"]
    )

    def apply_case(text):
        if case_option == "UPPERCASE":
            return str(text).upper()
        elif case_option == "Proper Case":
            return str(text).title()
        elif case_option == "lowercase":
            return str(text).lower()
        else:
            return str(text)

    # Step 6: Function to create DOCX (can be used for preview and final)
    def create_doc(limit_pages=None):
        doc = Document()
        total_records = len(df) if not limit_pages else min(len(df), limit_pages * records_per_page)

        for i in range(0, total_records, records_per_page):
            chunk = df.iloc[i:i + records_per_page]
            for _, row in chunk.iterrows():
                for col in selected_columns:
                    display_name = apply_case(column_rename_map.get(col, col))
                    value = apply_case(row[col]) if pd.notna(row[col]) else ""

                    p = doc.add_paragraph()
                    run = p.add_run(f"{display_name}: {value}")
                    run.font.size = Pt(font_size)

                    # Bold if in bold list
                    if column_rename_map[col] in bold_fields:
                        run.bold = True

                doc.add_paragraph()  # Spacer between records
            doc.add_page_break()

        output = BytesIO()
        doc.save(output)
        output.seek(0)
        return output

    # Step 7: Live Preview
    st.subheader("👀 Live Preview (First 2 Pages)")
    preview_output = create_doc(limit_pages=2)
    st.download_button(
        "📥 Download Preview (DOCX)",
        data=preview_output,
        file_name="Preview_VLE_Output.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    # Step 8: Final Download
    if st.button("🛠️ Generate Final DOCX"):
        final_output = create_doc()
        st.success("✅ Document generated successfully!")
        st.download_button(
            "📥 Download Final DOCX",
            data=final_output,
            file_name="VLE_Output.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
