import os
import tempfile
import inspect
import streamlit as st

from convert_pdf_to_docx import convert
from word_reindent import apply_template_bullets

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "Doc2.docx")

st.set_page_config(page_title="PDF → Notes", layout="centered")
st.title("PDF → Notes Generator")
st.write("Upload a PDF and download the formatted Word notes.")

pdf_file = st.file_uploader("Upload your PDF", type=["pdf"])

mode = st.radio(
    "PDF type",
    options=[
        "Slide deck (has bullets)",
        "Handout / no bullets (layout-based)",
    ],
    index=0,
)

no_bullets = mode.startswith("Handout")
run = st.button("Run", type="primary", disabled=(pdf_file is None))

if run:
    if not os.path.exists(TEMPLATE_PATH):
        st.error("Server misconfigured: Doc2.docx not found next to app.py.")
        st.stop()

    with st.spinner("Converting PDF → DOCX..."):
        with tempfile.TemporaryDirectory() as td:
            pdf_path = os.path.join(td, "input.pdf")
            intermediate_docx = os.path.join(td, "intermediate.docx")
            final_docx = os.path.join(td, "final.docx")

            # Save PDF
            with open(pdf_path, "wb") as f:
                f.write(pdf_file.getbuffer())

            # Script 1 (convert) — only pass flag if supported
            sig = inspect.signature(convert)
            params = sig.parameters
            if "no_bullets" in params:
                convert(pdf_path, intermediate_docx, no_bullets=no_bullets)
            elif "mode" in params:
                convert(pdf_path, intermediate_docx, mode=("handout" if no_bullets else "bullets"))
            else:
                if no_bullets:
                    st.warning(
                        "You selected 'Handout / no bullets', but convert() does not support that flag yet. "
                        "Running bullet mode."
                    )
                convert(pdf_path, intermediate_docx)

            # Read intermediate bytes BEFORE temp dir is deleted
            with open(intermediate_docx, "rb") as f:
                intermediate_bytes = f.read()

            # Script 2
            apply_template_bullets(intermediate_docx, TEMPLATE_PATH, final_docx)

            # IMPORTANT: read bytes BEFORE temp dir is deleted
            with open(final_docx, "rb") as f:
                final_bytes = f.read()

    st.success("Done.")

    st.download_button(
        "Download INTERMEDIATE DOCX",
        data=intermediate_bytes,
        file_name="notes_intermediate.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    st.download_button(
        "Download FINAL DOCX",
        data=final_bytes,
        file_name="notes.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
