import PyPDF2
import xlsxwriter

def extract_text_to_excel(pdf_file, attributes, excel_file):
    # Open the PDF file
    with open(pdf_file, 'rb') as file:
        pdf_reader = PyPDF2.PdfFileReader(file)

    # Create a new Excel file and add a worksheet
    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()

    # Define the column headers
    worksheet.write(0, 0, "Page Number")
    worksheet.write(0, 1, "Attribute")
    worksheet.write(0, 2, "Text")

    # Keep track of the row number in the Excel sheet
    row_num = 1

    # Iterate through each page of the PDF
    for page in range(pdf_reader.numPages):
        page_obj = pdf_reader.getPage(page)
        text = page_obj.extractText()

        # Check for the presence of the desired attributes in the text
        for attribute in attributes:
            if attribute in text:
                # Write the page number, attribute, and text to the Excel sheet
                worksheet.write(row_num, 0, page)
                worksheet.write(row_num, 1, attribute)
                worksheet.write(row_num, 2, text)
                row_num += 1

    # Save the Excel file
    workbook.close()

# Example usage
pdf_file = "sample.pdf"
attributes = ["John", "Doe", "address", "phone number"]
excel_file = "sample.xlsx"
extract_text_to_excel(pdf_file, attributes, excel_file)