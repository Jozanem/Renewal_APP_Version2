import os
import pythoncom
import win32com.client


def convert_docx_to_pdf(docx_path):
    """
    Converts a Word .docx file to PDF using Microsoft Word COM.
    """

    # --------------------------------------------------
    # ENSURE ABSOLUTE CLEAN WINDOWS PATH
    # --------------------------------------------------
    docx_path = os.path.abspath(docx_path)
    docx_path = os.path.normpath(docx_path)

    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"File not found: {docx_path}")

    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"

    pythoncom.CoInitialize()

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF
        doc.Close()

    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    return pdf_path
