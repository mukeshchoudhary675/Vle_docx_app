import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from io import BytesIO

st.title("📄 VLE Data Formatter (Dynamic Column Renaming)")

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
    font_size = st.slider("🔠 Font size", min_value=10, max_value=24, value=14)
    records_per_page = st.slider("📃 Records per page", min_value=1, max_value=5, value=1)

    # Step 5: Generate DOCX
    if st.button("🛠️ Generate DOCX"):
        doc = Document()

        for i in range(0, len(df), records_per_page):
            chunk = df.iloc[i:i + records_per_page]
            for _, row in chunk.iterrows():
                for col in selected_columns:
                    display_name = column_rename_map.get(col, col).upper()  # All caps
                    value = str(row[col]).upper() if pd.notna(row[col]) else ""  # All caps values too

                    p = doc.add_paragraph()
                    run = p.add_run(f"{display_name}: {value}")
                    run.font.size = Pt(font_size)

                    # Bold if in bold list
                    if display_name in [b.upper() for b in bold_fields]:
                        run.bold = True

                doc.add_paragraph()  # Spacer between records
            doc.add_page_break()

        # Save to memory
        output = BytesIO()
        doc.save(output)
        output.seek(0)

        st.success("✅ Document generated successfully!")
        st.download_button(
            "📥 Download DOCX",
            data=output,
            file_name="VLE_Output.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
