
import streamlit as st
import fitz
from docx import Document
from io import BytesIO

st.set_page_config(
    page_title="PDF to Word Converter",
    page_icon="📄",
    layout="centered"
)

st.title("📄 PDF to Word Converter")
st.write("Convert your PDF documents into editable Word files for free.")

uploaded_file = st.file_uploader(
    "Upload your PDF document",
    type=["pdf"]
)

if uploaded_file:
    if st.button("Convert to Word", type="primary"):
        try:
            pdf = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            document = Document()

            for page_number, page in enumerate(pdf):
                text = page.get_text("text")

                if text.strip():
                    if page_number > 0:
                        document.add_page_break()

                    for paragraph in text.split("\n"):
                        if paragraph.strip():
                            document.add_paragraph(paragraph)

            output = BytesIO()
            document.save(output)
            output.seek(0)

            st.success("Your Word document is ready!")

            st.download_button(
                label="⬇️ Download Word Document",
                data=output,
                file_name="converted_document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            if not any(page.get_text().strip() for page in pdf):
                st.warning(
                    "This PDF contains no extractable text. "
                    "Scanned PDF support will be added in a later step."
                )

        except Exception as error:
            st.error(f"Conversion failed: {error}")
