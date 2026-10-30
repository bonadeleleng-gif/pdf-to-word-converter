
import streamlit as st
import fitz
import pandas as pd
from docx import Document
from pptx import Presentation
from io import BytesIO
from PIL import Image
import pytesseract

st.title("📚 Universal Document Reader")
    page_title="Document Reader",
    page_icon="📚",
    layout="wide"
# Android PWA integration
import streamlit.components.v1 as components

components.html(
    """
    <script>
    const parentDoc = window.parent.document;

    if (!parentDoc.querySelector('link[rel="manifest"]')) {
        const manifest = parentDoc.createElement('link');
        manifest.rel = 'manifest';
        manifest.href = '/manifest.json';
        parentDoc.head.appendChild(manifest);
    }

    if (!parentDoc.querySelector('meta[name="theme-color"]')) {
        const theme = parentDoc.createElement('meta');
        theme.name = 'theme-color';
        theme.content = '#2563eb';
        parentDoc.head.appendChild(theme);
    }

    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('/service-worker.js')
            .catch(error => console.log('Service worker:', error));
    }
    </script>
    """,
    height=0,
)

)

st.title("📚 Universal Document Reader")
st.write("Upload, read, search and extract text from your documents.")

st.sidebar.title("About")
st.sidebar.write(
    "Read PDF, Word, Excel, PowerPoint, "
    "CSV, TXT and image files."
)

uploaded_files = st.file_uploader(
    "Upload your documents",
    type=[
        "pdf", "docx", "xlsx",
        "pptx", "txt", "csv",
        "png", "jpg", "jpeg"
    ],
    accept_multiple_files=True
)

def read_document(file):
    name = file.name.lower()
    data = file.getvalue()
    text = ""

    if name.endswith(".pdf"):
        pdf = fitz.open(stream=data, filetype="pdf")

        for i, page in enumerate(pdf):
            text += f"\n--- Page {i+1} ---\n"
            text += page.get_text("text")

    elif name.endswith(".docx"):
        doc = Document(BytesIO(data))

        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"

        for table in doc.tables:
            for row in table.rows:
                text += " | ".join(
                    cell.text for cell in row.cells
                ) + "\n"

    elif name.endswith(".xlsx"):
        excel = pd.ExcelFile(BytesIO(data))

        for sheet in excel.sheet_names:
            text += f"\n--- Sheet: {sheet} ---\n"
            df = pd.read_excel(
                excel, sheet_name=sheet, header=None
            )
            text += df.fillna("").to_string(
                index=False, header=False
            )

    elif name.endswith(".pptx"):
        presentation = Presentation(BytesIO(data))

        for i, slide in enumerate(presentation.slides):
            text += f"\n--- Slide {i+1} ---\n"

            for shape in slide.shapes:
                if shape.has_text_frame:
                    text += shape.text + "\n"

                if shape.has_table:
                    for row in shape.table.rows:
                        text += " | ".join(
                            cell.text for cell in row.cells
                        ) + "\n"

    elif name.endswith(".txt"):
        text = data.decode("utf-8", errors="replace")

    elif name.endswith(".csv"):
        df = pd.read_csv(BytesIO(data))
        text = df.fillna("").to_string(index=False)

    elif name.endswith((".png", ".jpg", ".jpeg")):
        image_file = Image.open(BytesIO(data))
        text = pytesseract.image_to_string(image_file)

    return text


if uploaded_files:
    st.success(
        f"{len(uploaded_files)} document(s) uploaded successfully."
    )

    for file in uploaded_files:
        with st.expander(f"📄 {file.name}", expanded=True):

            try:
                extracted_text = read_document(file)

                if extracted_text.strip():
                    st.subheader("Document contents")

                    st.text_area(
                        "Extracted text",
                        value=extracted_text,
                        height=350,
                        key=file.name
                    )

                    search = st.text_input(
                        "Search inside this document",
                        key="search_" + file.name
                    )

                    if search:
                        matches = [
                            line for line in extracted_text.splitlines()
                            if search.lower() in line.lower()
                        ]

                        st.write(f"Found {len(matches)} matching lines.")

                        for match in matches[:100]:
                            st.write("• " + match)

                    st.download_button(
                        "⬇️ Download extracted text",
                        data=extracted_text,
                        file_name=file.name.rsplit(".", 1)[0] + ".txt",
                        mime="text/plain",
                        key="download_" + file.name
                    )

                else:
                    st.warning(
                        "No readable text found in this document."
                    )

            except Exception as error:
                st.error(f"Could not read this file: {error}")

st.divider()
st.caption("Universal Document Reader | Free document extraction tool")
