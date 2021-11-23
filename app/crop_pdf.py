from PyPDF2 import PdfFileReader, PdfFileWriter
"""
Uses the PyPDF2 package to crop PDF files

"""

# Pass in a file name that will be in the local directory
def crop_pdf(file):

    # Initialize the reader object to the file
    reader = PdfFileReader(file, 'r')

    # Initialize the writer object
    writer = PdfFileWriter()

    # Iterate through all the pages in order to crop out the top bit
    for i in range(reader.getNumPages()):

        # Get individual page to crop
        page = reader.getPage(i)

        # Positions are set as Tuples with x,y coordinates. The module treats PDF pages like a coordinate plane or graph
        # EX: (30,89) where x = 30, and y = 89
        # For a more detailed explanation head to 
        # https://www.youtube.com/watch?v=K45PptGYHOU
        page.cropBox.setUpperLeft((0, page.cropBox.getUpperLeft()[1] - 25))
        page.cropBox.setLowerRight((page.cropBox.getUpperRight()[0], 0))

        # Add page to the writer object
        writer.addPage(page)

    # Open binary stream with chosen file name
    outstream = open(file[:11] + '-cropped.pdf', 'wb')

    # Write all the pages to the outstream object
    writer.write(outstream)

    # Close it
    outstream.close()