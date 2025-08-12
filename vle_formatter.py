import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Dynamic Excel Processor", layout="wide")
st.title("📊 Dynamic Excel Cleaner with Custom Headers + Audit Tracking")

uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    all_columns = df.columns.tolist()

    st.success(f"✅ Uploaded: {df.shape[0]} rows × {df.shape[1]} columns")

    # --- Main Output Selection ---
    st.subheader("📌 Select & Rename Columns for Final Output")
    selected_columns = st.multiselect(
        "Choose columns for the main cleaned file",
        options=all_columns,
        default=all_columns
    )

    rename_mapping = {}
    for col in selected_columns:
        new_name = st.text_input(f"Rename '{col}' to:", value=col)
        rename_mapping[col] = new_name

    # --- Audit Tracking Selection ---
    st.subheader("🛠️ Select & Rename Columns for Audit Tracking Sheet")
    audit_columns = st.multiselect(
        "Choose columns for Audit Tracking",
        options=all_columns
    )

    audit_rename_mapping = {}
    for col in audit_columns:
        new_name = st.text_input(f"Rename '{col}' to:", value=col, key=f"audit_{col}")
        audit_rename_mapping[col] = new_name

    # --- State Column ---
    state_column = st.selectbox(
        "🌍 Select 'State' column if available (or leave blank)",
        options=["None"] + all_columns,
        index=0
    )

    if st.button("🔄 Process and Generate Excel"):
        try:
            # Final output DF
            final_df = df[selected_columns].rename(columns=rename_mapping)

            # Audit Tracking DF
            audit_df = df[audit_columns].drop_duplicates().rename(columns=audit_rename_mapping)
            audit_df.insert(0, "Sr No.", range(1, len(audit_df) + 1))

            # Write to Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Audit Tracking sheet
                audit_df.to_excel(writer, sheet_name="Audit Tracking", index=False)

                # State-wise or single
                if state_column != "None":
                    for state, group in final_df.groupby(df[state_column]):
                        sheet_name = str(state)[:31] if pd.notna(state) else "Unknown"
                        group.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    final_df.to_excel(writer, sheet_name="Cleaned Data", index=False)

            st.success("✅ File processed successfully!")
            st.download_button(
                label="📥 Download Excel File",
                data=output.getvalue(),
                file_name="dynamic_output_with_custom_headers.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")

























# import streamlit as st
# import pandas as pd
# from docx import Document
# from docx.shared import Pt
# from io import BytesIO

# st.title("📄 VLE Data Formatter")

# uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])

# if uploaded_file:
#     df = pd.read_excel(uploaded_file)
#     st.success("Excel uploaded successfully!")
#     st.dataframe(df.head())

#     columns = st.multiselect("🧩 Select columns to include", df.columns.tolist(), default=df.columns.tolist()[:3])
#     font_size = st.slider("🔠 Font size", min_value=10, max_value=24, value=14)
#     records_per_page = st.slider("📃 Records per page", min_value=1, max_value=5, value=1)

#     if st.button("🛠️ Generate DOCX"):
#         doc = Document()
#         for i in range(0, len(df), records_per_page):
#             chunk = df.iloc[i:i+records_per_page]
#             for _, row in chunk.iterrows():
#                 for col in columns:
#                     p = doc.add_paragraph()
#                     run = p.add_run(f"{col.upper()}: {row[col]}")
#                     run.font.size = Pt(font_size)
#                 doc.add_paragraph()
#             doc.add_page_break()

#         output = BytesIO()
#         doc.save(output)
#         output.seek(0)

#         st.success("✅ Document generated!")
#         st.download_button("📥 Download DOCX", data=output, file_name="VLE_Output.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
