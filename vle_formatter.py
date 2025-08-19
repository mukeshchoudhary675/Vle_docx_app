import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt

st.title("📄 Excel → Word (TO/FROM Labels) Generator")

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("✅ File uploaded. Preview:")
    st.dataframe(df.head())

    # Select TO section columns
    to_columns = st.multiselect("Select TO Section Columns", df.columns)

    # Select FROM section column (only one)
    from_column = st.selectbox("Select FROM Section Column", df.columns)

    # Rename columns dynamically
    st.subheader("📝 Rename TO Columns")
    column_rename_map = {}
    for col in to_columns:
        new_name = st.text_input(f"Rename '{col}' as:", col)
        column_rename_map[col] = new_name

    # Case style options
    case_option = st.selectbox("Choose case style", ["UPPERCASE", "lowercase", "Proper Case"])

    # Font size options
    font_size_to = st.slider("Font size for TO section", 8, 30, 12)
    font_size_from = st.slider("Font size for FROM section", 8, 30, 10)

    # Bold fields selection
    bold_fields = st.multiselect("Select TO fields to make bold", list(column_rename_map.values()))

    # Extra blank line fields
    blank_line_fields = st.multiselect("Select TO fields after which a blank line should appear", list(column_rename_map.values()))

    # Apply case formatting function
    def apply_case(text):
        if case_option == "UPPERCASE":
            return str(text).upper()
        elif case_option == "lowercase":
            return str(text).lower()
        elif case_option == "Proper Case":
            return str(text).title()
        return str(text)

    def clean_value(val):
        # Prevent float PINCODE like 781128.0 → 781128
        if isinstance(val, float) and val.is_integer():
            return str(int(val))
        return str(val)

    # Function to generate Word doc
    def create_doc(df, to_columns, from_column):
        doc = Document()

        for _, row in df.iterrows():
            # --- TO Section ---
            p = doc.add_paragraph()
            run = p.add_run("TO")
            run.bold = True
            run.font.size = Pt(font_size_to)

            for col in to_columns:
                display_name = apply_case(column_rename_map.get(col, col))
                value = apply_case(clean_value(row[col])) if pd.notna(row[col]) else ""

                p = doc.add_paragraph()
                run = p.add_run(f"{display_name}: {value}")
                run.font.size = Pt(font_size_to)

                if display_name in bold_fields:
                    run.bold = True

                if display_name in blank_line_fields:
                    doc.add_paragraph()

            # --- FROM Section ---
            p = doc.add_paragraph()
            run = p.add_run("FROM")
            run.bold = True
            run.font.size = Pt(font_size_from)

            value = apply_case(clean_value(row[from_column])) if pd.notna(row[from_column]) else ""
            p = doc.add_paragraph()
            run = p.add_run(f"{apply_case(from_column)}: {value}")
            run.font.size = Pt(font_size_from)

            doc.add_page_break()

        return doc

    # Preview first 2 records in Streamlit
    st.subheader("👀 Preview (first 2 records)")
    for i, row in df.head(2).iterrows():
        st.markdown("### TO")
        for col in to_columns:
            display_name = apply_case(column_rename_map.get(col, col))
            value = apply_case(clean_value(row[col])) if pd.notna(row[col]) else ""
            if display_name in bold_fields:
                st.markdown(f"**{display_name}: {value}**")
            else:
                st.write(f"{display_name}: {value}")
            if display_name in blank_line_fields:
                st.text(" ")

        st.markdown("### FROM")
        value = apply_case(clean_value(row[from_column])) if pd.notna(row[from_column]) else ""
        st.write(f"{apply_case(from_column)}: {value}")
        st.write("---")

    # Download button
    if st.button("Generate Word File"):
        doc = create_doc(df, to_columns, from_column)
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        st.download_button("📥 Download Word File", buffer, "output.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
