import os
from comtypes.client import CreateObject

def docx_to_pdf(docx_path, pdf_path):
    word = CreateObject("Word.Application")
    doc = word.Documents.Open(docx_path)
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    word.Quit()

directory = "." #aktualny katalog  
for filename in os.listdir(directory):
    if filename.endswith(".docx"):
        docx_path = os.path.join(directory, filename)
        pdf_path = os.path.join(directory, filename[:-5] + ".pdf")
        docx_to_pdf(docx_path, pdf_path)
