import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt

# -------- Helper Functions --------
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
        return str(int(val))  # fix PINCODE 781128.0 → 781128
    return str(val)

# -------- DOCX Creator --------
def create_doc(df, from_column, to_columns, column_rename_map, case_style, from_font_size, to_font_size, bold_fields, blankline_fields):
    doc = Document()
    for _, row in df.iterrows():
        # -------- FROM Section --------
        doc.add_paragraph("FROM:")
        if from_column:
            from_text = clean_value(row[from_column]) if pd.notna(row[from_column]) else ""
            from_text = apply_case(from_text, case_style)
            p = doc.add_paragraph()
            run = p.add_run(from_text)
            run.font.size = Pt(from_font_size)

        doc.add_paragraph()  # spacing between FROM and TO

        # -------- TO Section --------
        doc.add_paragraph("TO:")
        for col in to_columns:
            display_name = apply_case(column_rename_map.get(col, col), case_style)
            value = apply_case(clean_value(row[col]), case_style) if pd.notna(row[col]) else ""

            p = doc.add_paragraph()
            run = p.add_run(f"{display_name}: {value}")
            run.font.size = Pt(to_font_size)

            # Bold option
            if column_rename_map[col] in bold_fields:
                run.bold = True

            # Blank line option
            if column_rename_map[col] in blankline_fields:
                doc.add_paragraph()

        doc.add_page_break()
    return doc

# -------- Streamlit App --------
st.title("📄 Excel to Multi-Page Word Generator (FROM / TO Format)")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"File loaded with {df.shape[0]} rows and {df.shape[1]} columns.")

    # FROM Column
    st.subheader("FROM Section")
    from_column = st.selectbox("Select FROM Address Column", [""] + df.columns.tolist())
    from_font_size = st.slider("Font Size for FROM Section", 8, 30, 12)

    # TO Columns
    st.subheader("TO Section")
    to_columns = st.multiselect("Select TO Columns", df.columns.tolist())

    if to_columns:
        # Rename Columns
        st.subheader("Rename TO Columns for Output")
        column_rename_map = {}
        for col in to_columns:
            column_rename_map[col] = st.text_input(f"Rename '{col}' as:", value=col)

        # Case Style
        case_style = st.radio("Select Case Style", ["UPPERCASE", "lowercase", "Proper Case"])

        # Font Size for TO
        to_font_size = st.slider("Font Size for TO Section", 8, 30, 12)

        # Bold Fields
        bold_fields = st.multiselect("Select TO Fields to Make Bold", list(column_rename_map.values()))

        # Blank Line After Fields
        blankline_fields = st.multiselect("Select TO Fields After Which to Add Blank Line", list(column_rename_map.values()))

        # -------- Preview --------
        st.subheader("🔍 Preview (First 2 Records)")
        preview_df = df[[from_column] + to_columns].head(2) if from_column else df[to_columns].head(2)
        for _, row in preview_df.iterrows():
            if from_column:
                st.markdown("**FROM:**")
                st.write(apply_case(clean_value(row[from_column]), case_style))

            st.markdown("**TO:**")
            for col in to_columns:
                display_name = apply_case(column_rename_map[col], case_style)
                value = apply_case(clean_value(row[col]), case_style) if pd.notna(row[col]) else ""

                if column_rename_map[col] in bold_fields:
                    st.markdown(f"**{display_name}: {value}**")
                else:
                    st.write(f"{display_name}: {value}")

                if column_rename_map[col] in blankline_fields:
                    st.text(" ")

            st.markdown("---")

        # -------- Generate DOCX --------
        if st.button("Generate Word File"):
            doc = create_doc(
                df, from_column, to_columns,
                column_rename_map, case_style,
                from_font_size, to_font_size,
                bold_fields, blankline_fields
            )

            output = BytesIO()
            doc.save(output)
            output.seek(0)

            st.download_button(
                label="📥 Download Word File",
                data=output,
                file_name="output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
