import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt

st.title("📄 Excel → Word (TO first, then FROM) — Dynamic Formatting")

# ---------- Helpers ----------
def apply_case(text: str, mode: str) -> str:
    s = "" if text is None else str(text)
    if mode == "UPPERCASE":
        return s.upper()
    if mode == "lowercase":
        return s.lower()
    if mode == "Proper Case":
        return s.title()
    return s  # Original

def clean_value(val) -> str:
    # Handles 781128.0 -> 781128 and trims strings
    if val is None:
        return ""
    try:
        if isinstance(val, float) and float(val).is_integer():
            return str(int(val))
    except Exception:
        pass
    return str(val).strip()

# Build a paragraph line in DOCX with dynamic bold & size
def add_line(doc, text: str, font_size_pt: int, bold: bool = False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(font_size_pt)
    run.bold = bold
    return p

# ---------- App ----------
uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Read as strings to preserve IDs/phone/pincode exactly
    df = pd.read_excel(uploaded_file, dtype=str)
    st.success(f"Loaded {df.shape[0]} rows × {df.shape[1]} columns")
    st.expander("Preview first rows (raw)").dataframe(df.head())

    # --- Column selection ---
    st.subheader("TO Section (appears first)")
    to_columns = st.multiselect("Choose TO fields (order matters)", df.columns.tolist())

    st.subheader("FROM Section (appears below TO)")
    from_column = st.selectbox("Choose FROM field (single column)", ["(none)"] + df.columns.tolist())
    use_from = from_column != "(none)"

    # --- Renaming TO fields (user labels) ---
    st.subheader("Rename TO Fields (labels shown in output)")
    column_rename_map = {}
    for col in to_columns:
        column_rename_map[col] = st.text_input(f"Rename '{col}' →", value=col, key=f"rename_{col}")

    # --- Text casing for ALL labels & values ---
    st.subheader("Text Casing")
    case_option = st.selectbox("Apply casing to labels & values", ["Original", "UPPERCASE", "Proper Case", "lowercase"])

    # --- Font sizes (4 sliders) ---
    st.subheader("Font Sizes")
    to_label_size = st.slider("TO label font size (e.g., 'TO:')", 8, 36, 14)
    to_data_size  = st.slider("TO lines font size (e.g., 'NAME: MUKESH')", 8, 36, 14)
    from_label_size = st.slider("FROM label font size (e.g., 'FROM:')", 8, 36, 12)
    from_data_size  = st.slider("FROM line font size (address line)", 8, 36, 12)

    # --- Bold controls ---
    st.subheader("Bold Settings")
    bold_to_label   = st.checkbox("Make 'TO:' label bold", value=True)
    bold_from_label = st.checkbox("Make 'FROM:' label bold", value=True)
    # Multiselect uses RAW renamed labels (before casing) to avoid mismatch
    to_labels_list = [column_rename_map[c] for c in to_columns]
    bold_to_fields = st.multiselect("Make these TO lines bold", to_labels_list)
    blankline_after = st.multiselect("Add a blank line after these TO lines", to_labels_list)
    bold_from_line = st.checkbox("Make FROM address line bold", value=False)

    # ---------- Preview (first 2 records) ----------
    st.subheader("👀 Live Preview (first 2 records)")
    preview_rows = min(2, len(df))
    if preview_rows == 0:
        st.info("No data to preview.")
    else:
        for i in range(preview_rows):
            row = df.iloc[i]

            # TO label
            to_label_disp = apply_case("TO:", case_option)
            st.markdown(
                f"<div style='font-size:{to_label_size}px; font-weight:{'700' if bold_to_label else '400'}'>{to_label_disp}</div>",
                unsafe_allow_html=True
            )
            # TO lines
            for col in to_columns:
                raw_label = column_rename_map[col]                     # unchanged (user-entered) label
                show_label = apply_case(raw_label, case_option)        # cased for display
                value = apply_case(clean_value(row.get(col, "")), case_option)

                bold_line = raw_label in bold_to_fields                # compare with RAW label (fixes earlier bug)
                line_html = f"<div style='font-size:{to_data_size}px; font-weight:{'700' if bold_line else '400'}'>{show_label}: {value}</div>"
                st.markdown(line_html, unsafe_allow_html=True)

                if raw_label in blankline_after:
                    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)  # small blank gap

            # FROM section (optional)
            if use_from:
                from_label_disp = apply_case("FROM:", case_option)
                st.markdown(
                    f"<div style='margin-top:8px; font-size:{from_label_size}px; font-weight:{'700' if bold_from_label else '400'}'>{from_label_disp}</div>",
                    unsafe_allow_html=True
                )
                from_value = apply_case(clean_value(row.get(from_column, "")), case_option)
                st.markdown(
                    f"<div style='font-size:{from_data_size}px; font-weight:{'700' if bold_from_line else '400'}'>{from_value}</div>",
                    unsafe_allow_html=True
                )

            st.markdown("<hr>", unsafe_allow_html=True)

    # ---------- Create DOCX ----------
    def build_doc(full_df: pd.DataFrame) -> BytesIO:
        doc = Document()

        for _, row in full_df.iterrows():
            # TO label
            add_line(doc, apply_case("TO:", case_option), to_label_size, bold=bold_to_label)

            # TO lines
            for col in to_columns:
                raw_label = column_rename_map[col]  # unchanged label
                show_label = apply_case(raw_label, case_option)
                value = apply_case(clean_value(row.get(col, "")), case_option)

                # Entire line bold if user selected the RAW label
                make_bold = raw_label in bold_to_fields
                add_line(doc, f"{show_label}: {value}", to_data_size, bold=make_bold)

                # Blank line after certain fields
                if raw_label in blankline_after:
                    doc.add_paragraph()

            # FROM section (optional)
            if use_from:
                add_line(doc, apply_case("FROM:", case_option), from_label_size, bold=bold_from_label)
                from_value = apply_case(clean_value(row.get(from_column, "")), case_option)
                add_line(doc, from_value, from_data_size, bold=bold_from_line)

            # Page break per record
            doc.add_page_break()

            # (Optional) remove the last page break later if needed

        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    # ---------- Download ----------
    if st.button("Generate Word (.docx)"):
        output = build_doc(df)
        st.success("✅ DOCX generated")
        st.download_button(
            "📥 Download DOCX",
            data=output,
            file_name="output.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
