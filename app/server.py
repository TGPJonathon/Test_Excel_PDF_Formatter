from flask import Flask, request, send_file, Response, g
from flask_cors import CORS
from excel_formatter import make_excel_sheet
from crop_pdf import crop_pdf

import jpype
import jdk
import asposecells
import os

"""
Goal of this API is to send back a formatted PDF or an Excel file.
Based on contents of POST request sent from UI
    pdf parameter = True => You get PDF
    pdf parameter = False => You get Excel/xlsx
"""

# Create Flask App
app = Flask(__name__)

# Used to get around CORS for browsers
CORS(app)


# Only Runs once, when the server is started
@app.before_first_request
def before_first_request_func():

##################################################################
#          Install Java Since it's Needed to Convert Excel to PDF
##################################################################

  # Check to see if Java is installed
  is_java_installed = False

  # Will iterate through immediate directory to check
  for dir in os.listdir():
    if(dir.startswith("jdk")):
      is_java_installed = True

  # If not there, install it
  if(not is_java_installed):
    jdk.install('11', jre=True, path="./")

  #Set location of Java installation
  os.environ["JAVA_HOME"] = "./"

  #Start Java Virtual Machine
  jpype.startJVM()



#####################################################
#      End Java Install
#####################################################


@app.route("/", methods=["POST"])
def send_excel():

#####################################################
#      Parameters determine format of Excel sheet
#####################################################
  
  excel_data = request.get_json().get('data')

  # g module allows variables to accessed globally within the same response
  # Doesn't function like a normal Global
  g.title = request.get_json().get('excelFormattingParameters')[0]['fileTitle']
  highlight_rows = request.get_json().get('excelFormattingParameters')[0]['highlight']
  column_widths = request.get_json().get('excelFormattingParameters')[0]['widths']
  excel_sheet_title = request.get_json().get('excelFormattingParameters')[0]['excelTitle']
  pdf_or_no = request.get_json().get('excelFormattingParameters')[0]['pdf']

  output = make_excel_sheet(g.get('title'), excel_data, highlight_rows, column_widths, excel_sheet_title, pdf_or_no)

#####################################################
#      End
#####################################################


#####################################################
#      Start Excel Conversion to PDF
#####################################################

  #Package only works with Java
  from asposecells.api import SaveFormat, Workbook   

  # If the user wants a PDF the make_excel_sheet function will return 0 as a value
  # make_excel_sheet will create a xlsx file in the local directory in that case
  if(output == 0):
    
    #asposeCells package will open the created excel sheet based on the title
    workbook = Workbook(g.get('title') + ".xlsx")

    #Converts the excel book to a PDF
    workbook.save("xlsx-to-pdf.pdf", SaveFormat.PDF)

    #aspose_cells package puts a red banner at the top of each Excel page
    #this function will crop it out of each page
    crop_pdf("xlsx-to-pdf.pdf")

#####################################################
#      End
#####################################################



  #Sends the file as a response to the browser to be downloaded
  #download_name has to be set even though it won't be used, the browser will normally handle the name of the file
  return send_file(output if output != 0 else "xlsx-to-pdf-cropped.pdf", download_name="not_needed.pdf" if output == 0 else "not_needed.xlsx", as_attachment=True)


@app.after_request
def after_request_func(response):

    """ 
    This function will run after a request, as long as no exceptions occur.
    It must take and return the same parameter - an instance of response_class.
    This is a good place to do some application cleanup.
    """

    close = Response("Testing")

    # Will perform some clean up of local files
    # Cannot remove the xlsx-to-pdf-cropped.pdf file because the response hasn't fully ended so the file wouldn't be sent
    # Might be okay though, since the file gets overwritten each api call
    def process_after_request():

      if os.path.isfile("xlsx-to-pdf.pdf"):
        os.remove("xlsx-to-pdf.pdf")
      
      if(g.get("title")):
        if os.path.isfile(g.get("title") + ".xlsx"):
          os.remove(g.get("title") + ".xlsx")

    # Will call function at end of request
    close.call_on_close(process_after_request())

    return response



if __name__ == '__main__':
  app.run()