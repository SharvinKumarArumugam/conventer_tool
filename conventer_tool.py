import streamlit as st
from pdf2docx import Converter
from docx import Document
from reportlab.pdfgen import canvas
import img2pdf
from io import BytesIO
import os
import tempfile

st.set_page_config(page_title="Document Converter", layout="centered")

st.title("📄 Document Conversion Tool")

menu = st.sidebar.selectbox(
    "Choose a conversion type:",
    ["Select Option", "PDF to Word", "Image to PDF", "Word to PDF"]
)

def save_temp_file(uploaded_file):
    temp_file = tempfile.NamedTemporaryFile(delete=False)
    temp_file.write(uploaded_file.read())
    temp_file_path = temp_file.name
    temp_file.close()
    return temp_file_path

if menu == "PDF to Word":
    st.header("Convert PDF to Word")
    uploaded_pdf = st.file_uploader("Upload a PDF file", type=["pdf"])
    
    if uploaded_pdf and st.button("Convert"):
        pdf_path = save_temp_file(uploaded_pdf)
        word_output = os.path.splitext(uploaded_pdf.name)[0] + ".docx"
        word_path = os.path.join(tempfile.gettempdir(), word_output)

        cv = Converter(pdf_path)
        cv.convert(word_path, start=0, end=None)
        cv.close()

        with open(word_path, "rb") as f:
            st.success("Conversion successful!")
            st.download_button("Download Word File", f.read(), file_name=word_output)

elif menu == "Image to PDF":
    st.header("Convert Image to PDF")
    uploaded_image = st.file_uploader("Upload an image", type=["jpg", "jpeg", "png"])
    
    if uploaded_image and st.button("Convert"):
        image_path = save_temp_file(uploaded_image)
        output_pdf = os.path.splitext(uploaded_image.name)[0] + ".pdf"
        pdf_path = os.path.join(tempfile.gettempdir(), output_pdf)

        with open(pdf_path, "wb") as f:
            f.write(img2pdf.convert(image_path))
        
        with open(pdf_path, "rb") as f:
            st.success("Image successfully converted to PDF!")
            st.download_button("Download PDF", f.read(), file_name=output_pdf)

elif menu == "Word to PDF":
    st.header("Convert Word to PDF")
    uploaded_docx = st.file_uploader("Upload a Word (.docx) file", type=["docx"])
    
    if uploaded_docx and st.button("Convert"):
        docx_path = save_temp_file(uploaded_docx)
        pdf_output = os.path.splitext(uploaded_docx.name)[0] + ".pdf"
        pdf_path = os.path.join(tempfile.gettempdir(), pdf_output)

        doc = Document(docx_path)
        c = canvas.Canvas(pdf_path)
        width, height = c._pagesize
        y = height - 40

        for para in doc.paragraphs:
            text = para.text
            c.drawString(40, y, text)
            y -= 15
            if y <= 40:
                c.showPage()
                y = height - 40
        c.save()

        with open(pdf_path, "rb") as f:
            st.success("Word successfully converted to PDF!")
            st.download_button("Download PDF", f.read(), file_name
