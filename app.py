import os
import tempfile
import streamlit as st

from convert_pdf_to_docx import convert
from word_reindent import apply_template_bullets

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "Doc2.docx")

st.set_page_config(page_title="PDF → Notes", layout="centered")
st.title("PDF → Notes Generator")
st.write("Upload a PDF and download the formatted Word notes.")

pdf_file = st.file_uploader("Upload your PDF", type=["pdf"])

# Conversion mode (keeps old behaviour by default)
mode_label = st.radio(
    "Conversion mode",
    options=[
        "Bullets only (original)",
        "All lines as bullets (indent-based)",
    ],
    index=0,
    disabled=(pdf_file is None),
)

mode = "bullets_only" if mode_label.startswith("Bullets only") else "all_lines"

run = st.button("Run", type="primary", disabled=(pdf_file is None))

if run:
    if not os.path.exists(TEMPLATE_PATH):
        st.error("Server misconfigured: Doc2.docx not found.")
        st.stop()

    with st.spinner("Converting PDF → DOCX..."):
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

            st.success("Done.")

            col1, col2 = st.columns(2)

            with col1:
                with open(intermediate_docx, "rb") as f:
                    st.download_button(
                        "Download intermediate DOCX",
                        f,
                        file_name="notes_intermediate.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )

            with col2:
                with open(final_docx, "rb") as f:
                    st.download_button(
                        "Download final DOCX",
                        f,
                        file_name="notes.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
