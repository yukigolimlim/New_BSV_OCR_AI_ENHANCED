"""
main.py — DocExtract Pro
=========================
Entry point. Run this file to launch the application:

    python main.py

Dependencies (install with pip):
    customtkinter, pillow, paddleocr, groq, pdfplumber, pdf2image,
    openpyxl, pandas, python-docx, opencv-python, python-dotenv
"""
from dotenv import load_dotenv
load_dotenv()

from app import DocExtractorApp

if __name__ == "__main__":
    app = DocExtractorApp()
    app.mainloop()