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
    help=(
        "Choose 'Handout / no bullets' for PDFs that are mostly continuous text with headings, "
        "tables, or paragraphs (few/no bullet glyphs)."
    ),
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

            # ---- Script 1 (convert) ----
            # Backwards-compatible: only pass the flag if convert() supports it.
            try:
                sig = inspect.signature(convert)
                params = sig.parameters

                if "no_bullets" in params:
                    convert(pdf_path, intermediate_docx, no_bullets=no_bullets)
                elif "mode" in params:
                    convert(pdf_path, intermediate_docx, mode=("handout" if no_bullets else "bullets"))
                else:
                    # convert() doesn't accept a mode flag yet
                    if no_bullets:
                        st.warning(
                            "You selected 'Handout / no bullets', but convert() does not yet support a mode flag. "
                            "It will run in bullet mode for now."
                        )
                    convert(pdf_path, intermediate_docx)

            except Exception as e:
                st.error(f"Conversion failed: {e}")
                st.stop()

            # ---- Script 2 (template bullets) ----
            try:
                apply_template_bullets(intermediate_docx, TEMPLATE_PATH, final_docx)
            except Exception as e:
                st.error(f"Template formatting failed: {e}")
                st.stop()

    st.success("Done.")
    with open(final_docx, "rb") as f:
        st.download_button(
            "Download DOCX",
            f,
            file_name="notes.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
