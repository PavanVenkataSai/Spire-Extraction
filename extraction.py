# from spire.pdf.common import *
# from spire.pdf import *

# # Create a PdfDocument object
# doc = PdfDocument()

# all_text = ""
# doc.LoadFromFile(r"C:\Users\CVHS\pavan\Alternative\spire.ai\Carls_Digital (1).pdf")


# for page_number in range(doc.Pages.Count):
#     page = doc.Pages[page_number]
#     textExtractor = PdfTextExtractor(page)
#     extractOptions = PdfTextExtractOptions()
#     extractOptions.IsExtractAllText = True
#     text = textExtractor.ExtractText(extractOptions)

#     # Append the extracted text to all_text
#     all_text += text + "\n"  # Add a newline to separate pages

# # Optionally print the text from each file
# print(all_text)

# with open('output.txt', 'w', encoding='utf-8') as file:
#     file.write(all_text)

from spire.pdf.common import *
from spire.pdf import *

# Create a PdfDocument object
pdf = PdfDocument()
# Load a PDF document
pdf.LoadFromFile(r"C:\Users\CVHS\pavan\Alternative\spire.ai\Carls_Digital (1).pdf")

# Create an XlsxLineLayoutOptions object to specify the conversion options
# Parameters: convertToMultipleSheet, rotatedText, splitCell, wrapText, overlapText
convertOptions = XlsxLineLayoutOptions(True, True, False, True, False)

# Set the conversion options
pdf.ConvertOptions.SetPdfToXlsxOptions(convertOptions)

# Save the PDF document to Excel XLSX format
pdf.SaveToFile("PdfToExcel.xlsx", FileFormat.XLSX)
pdf.Close()
