# from spire.doc import *
# from spire.doc.common import *

# # Create word document
# document = Document()

# # Load a doc or docx file
# document.LoadFromFile(r"C:\Users\CVHS\Downloads\PATTERSON_Profile Doc.docx")

# #Save the document to PDF
# document.SaveToFile(r"ToPDF.pdf", FileFormat.PDF)
# document.Close()

from docx2pdf import convert

convert(r"C:\Users\CVHS\Downloads\PATTERSON_Profile Doc.docx", r"ToPDF.pdf")