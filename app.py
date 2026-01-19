import os
import tempfile
import streamlit as st

from convert_pdf_to_docx import convert
from word_reindent import apply_template_bullets

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "Doc2.docx")

st.set_page_config(page_title="PDF → Notes", layout="centered")

# ----------------------------
# Header
# ----------------------------
st.title("PDF → Notes Generator")
st.caption("Upload a lecture slides PDF and get clean Word notes in Doc2.docx style.")

# ----------------------------
# Help / Docs
# ----------------------------
with st.expander("How to use", expanded=True):
    st.markdown(
        """
1. **Upload your PDF** (exported lecture slides).
2. Choose a **conversion mode**:
   - **Bullets only (original):** Best when the slides use bullet symbols (➣, •, etc).
   - **All lines as bullets:** Best when slides are mostly lines of text without bullet symbols.
3. Click **Run**.
4. Download:
   - **Intermediate DOCX** (raw structure, useful for debugging)
   - **Final DOCX** (formatted with Word-native bullets using Doc2.docx)
"""
    )

with st.expander("What’s happening under the hood? (high-level)", expanded=False):
    st.markdown(
        """
### `convert_pdf_to_docx.py`
Extracts text from the PDF and creates an **intermediate DOCX**:
- Detects slide titles/headings
- Converts content into bullets
- Uses x-position (coordinates) to infer indentation levels
- Removes obvious footer noise

### `word_reindent.py`
Takes the intermediate DOCX and rewrites it into a **final DOCX** that uses:
- The exact Word-native bullet/numbering style from **Doc2.docx**
- Consistent font (Aptos 12)
- Clean formatting (no weird spacing, no bold carry-over)

### `Doc2.docx`
Your **style template**:
- Defines the bullet formatting you want (so bullets behave like normal Word lists)
"""
    )

st.divider()

# ----------------------------
# Inputs
# ----------------------------
st.subheader("1) Upload")
pdf_file = st.file_uploader("Upload your PDF", type=["pdf"], label_visibility="collapsed")

st.subheader("2) Choose mode")
mode_label = st.radio(
    "Conversion mode",
    options=[
        "Bullets only (original)",
        "All lines as bullets (indent-based)",
    ],
    index=0,
    disabled=(pdf_file is None),
    help=(
        "Use Bullets only if the slides have bullet symbols (➣, •). "
        "Use All lines if the slides are mostly plain lines of text."
    ),
)

mode = "bullets_only" if mode_label.startswith("Bullets only") else "all_lines"

st.subheader("3) Run")

run = st.button(
    "Run conversion",
    type="primary",
    disabled=(pdf_file is None),
    use_container_width=True,
)

# ----------------------------
# Run pipeline
# ----------------------------
if run:
    if not os.path.exists(TEMPLATE_PATH):
        st.error("Server misconfigured: Doc2.docx not found next to app.py.")
        st.stop()

    # Small status block (nice UX)
    status = st.status("Working...", expanded=True)
    status.write("Saving uploaded PDF…")
    status.write("Converting PDF → intermediate DOCX…")
    status.write("Applying Doc2.docx bullet formatting → final DOCX…")

    with tempfile.TemporaryDirectory() as td:
        pdf_path = os.path.join(td, "input.pdf")
        intermediate_docx = os.path.join(td, "intermediate.docx")
        final_docx = os.path.join(td, "final.docx")

        # Save PDF
        with open(pdf_path, "wb") as f:
            f.write(pdf_file.getbuffer())

        # Script 1
        convert(pdf_path, intermediate_docx, mode=mode)

        # Script 2
        apply_template_bullets(intermediate_docx, TEMPLATE_PATH, final_docx)

        status.update(label="Done ✅", state="complete", expanded=False)
        st.success("Conversion complete. Download your files below.")

        st.divider()
        st.subheader("Downloads")

        col1, col2 = st.columns(2)

        with col1:
            with st.container(border=True):
                st.markdown("**Intermediate DOCX**")
                st.caption("Raw structure output (useful for debugging).")
                with open(intermediate_docx, "rb") as f:
                    st.download_button(
                        "Download intermediate",
                        f,
                        file_name="notes_intermediate.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

        with col2:
            with st.container(border=True):
                st.markdown("**Final DOCX**")
                st.caption("Formatted using Doc2.docx (Word-native bullets).")
                with open(final_docx, "rb") as f:
                    st.download_button(
                        "Download final notes",
                        f,
                        file_name="notes.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
