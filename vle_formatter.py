import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt

# ----------- Helper Functions -------------
def apply_case(text, case_style):
    if case_style == "UPPERCASE":
        return str(text).upper()
    elif case_style == "lowercase":
        return str(text).lower()
    elif case_style == "Proper Case":
        return str(text).title()
    else:
        return str(text)

def clean_value(val):
    if isinstance(val, float) and val.is_integer():
        return str(int(val))   # fix PINCODE like 781128.0 → 781128
    return str(val)

# ----------- DOCX Creation ----------------
def create_doc(df, selected_columns, column_rename_map, case_style, font_size, bold_fields, blankline_fields):
    doc = Document()
    for _, row in df.iterrows():
        for col in selected_columns:
            display_name = apply_case(column_rename_map.get(col, col), case_style)
            value = apply_case(clean_value(row[col]), case_style) if pd.notna(row[col]) else ""

            p = doc.add_paragraph()
            run = p.add_run(f"{display_name}: {value}")
            run.font.size = Pt(font_size)

            # Bold if chosen
            if column_rename_map[col] in bold_fields:
                run.bold = True

            # Add blank line if selected
            if column_rename_map[col] in blankline_fields:
                doc.add_paragraph()

        doc.add_page_break()
    return doc

# ----------- Streamlit App ----------------
st.title("📄 Excel to Multi-Page Word Generator")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"File loaded with {df.shape[0]} rows and {df.shape[1]} columns.")

    # Column Selection
    selected_columns = st.multiselect("Select Columns to Include", df.columns.tolist())

    if selected_columns:
        # Rename Columns
        st.subheader("Rename Columns for Output")
        column_rename_map = {}
        for col in selected_columns:
            column_rename_map[col] = st.text_input(f"Rename '{col}' as:", value=col)

        # Case Style
        case_style = st.radio("Select Case Style", ["UPPERCASE", "lowercase", "Proper Case"])

        # Font Size
        font_size = st.slider("Font Size", 8, 30, 12)

        # Bold Fields
        bold_fields = st.multiselect("Select Fields to Make Bold", list(column_rename_map.values()))

        # Blank Line After Fields
        blankline_fields = st.multiselect("Select Fields After Which to Add Blank Line", list(column_rename_map.values()))

        # Preview First 2 Records
        st.subheader("🔍 Preview (First 2 Records)")
        preview_df = df[selected_columns].head(2)
        for _, row in preview_df.iterrows():
            for col in selected_columns:
                display_name = apply_case(column_rename_map[col], case_style)
                value = apply_case(clean_value(row[col]), case_style) if pd.notna(row[col]) else ""

                if column_rename_map[col] in bold_fields:
                    st.markdown(f"**{display_name}: {value}**")
                else:
                    st.write(f"{display_name}: {value}")

                if column_rename_map[col] in blankline_fields:
                    st.text(" ")

            st.markdown("---")

        # Generate DOCX
        if st.button("Generate Word File"):
            doc = create_doc(df, selected_columns, column_rename_map, case_style, font_size, bold_fields, blankline_fields)

            output = BytesIO()
            doc.save(output)
            output.seek(0)

            st.download_button(
                label="📥 Download Word File",
                data=output,
                file_name="output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
