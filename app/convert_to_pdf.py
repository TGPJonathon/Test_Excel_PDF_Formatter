import asposecells
from io import BytesIO
from crop_pdf import crop_pdf

def convert_to_pdf(title):
    #Package only works with Java
    from asposecells.api import SaveFormat, Workbook   
      
    #asposeCells package will open the created excel sheet based on the title
    workbook = Workbook(title + ".xlsx")

    #Converts the excel book to a PDF
    workbook.save("xlsx-to-pdf.pdf", SaveFormat.PDF)

    #aspose_cells package puts a red banner at the top of each Excel page
    #this function will crop it out of each page
    crop_pdf("xlsx-to-pdf.pdf")

    return_data = BytesIO()
    with open("xlsx-to-pdf-cropped.pdf", 'rb') as fo:
        return_data.write(fo.read())

    return_data.seek(0)

    return return_data