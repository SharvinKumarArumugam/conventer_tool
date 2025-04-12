from pdf2docx import Converter
from docx import Document
from reportlab.pdfgen import canvas
import img2pdf
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

# Setup Tkinter root
root = tk.Tk()
root.withdraw()  # Hide the main window

# Function to convert PDF to Word
def convert_pdf_to_word():
    pdf_file = filedialog.askopenfilename(title="Select PDF file", filetypes=[("PDF files", "*.pdf")])
    if not pdf_file:
        return
    word_file = filedialog.asksaveasfilename(title="Save Word file as", defaultextension=".docx", filetypes=[("Word files", "*.docx")])
    if not word_file:
        return
    cv = Converter(pdf_file)
    cv.convert(word_file, start=0, end=None)
    cv.close()
    messagebox.showinfo("Success", f"PDF converted to Word:\n{word_file}")

# Function to convert Image to PDF
def convert_image_to_pdf():
    image_path = filedialog.askopenfilename(title="Select image file", filetypes=[("Image files", "*.jpg *.jpeg *.png")])
    if not image_path:
        return
    output_pdf = filedialog.asksaveasfilename(title="Save PDF as", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if not output_pdf:
        return
    with open(output_pdf, "wb") as f:
        f.write(img2pdf.convert(image_path))
    messagebox.showinfo("Success", f"Image converted to PDF:\n{output_pdf}")

# Function to convert Word to PDF
def convert_word_to_pdf():
    docx_file = filedialog.askopenfilename(title="Select Word file", filetypes=[("Word files", "*.docx")])
    if not docx_file:
        return
    pdf_file = filedialog.asksaveasfilename(title="Save PDF as", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if not pdf_file:
        return
    doc = Document(docx_file)
    c = canvas.Canvas(pdf_file)
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
    messagebox.showinfo("Success", f"Word converted to PDF:\n{pdf_file}")

# Main menu loop
def main():
    while True:
        choice = simpledialog.askinteger(
            "Document Conversion Menu",
            "Select a service:\n1. PDF to Word\n2. Image to PDF\n3. Word to PDF\n4. Exit"
        )

        if choice == 1:
            convert_pdf_to_word()
        elif choice == 2:
            convert_image_to_pdf()
        elif choice == 3:
            convert_word_to_pdf()
        elif choice == 4:
            break
        else:
            messagebox.showwarning("Invalid choice", "Please select a valid option (1-4).")

if __name__ == "__main__":
    main()
