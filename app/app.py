from flask import Flask, request, send_file, Response, g
from convert_to_pdf import convert_to_pdf
from flask_healthz import healthz
from flask_cors import CORS
from excel_formatter import make_excel_sheet
from io import BytesIO
from convert_to_pdf import convert_to_pdf
from readiness_liveness import liveness, readiness

import jpype
import jdk
import os

"""
Goal of this API is to send back a formatted PDF or an Excel file.
Based on contents of POST request sent from UI
    pdf parameter = True => You get PDF
    pdf parameter = False => You get Excel/xlsx
"""

# Create Flask App
app = Flask(__name__)
app.register_blueprint(healthz, url_prefix='/healthz')

# Used to get around CORS for browsers
CORS(app)


# Only Runs once, when the server is started
@app.before_first_request
def before_first_request_func():

##################################################################
#          Install Java Since it's Needed to Convert Excel to PDF

#          *********IMPORTANT************
#           Only for Local Testing Purposes
#           Won't need when Containerized
##################################################################

  # # Check to see if Java is installed
  is_java_installed = False

  # # Will iterate through immediate directory to check
  for dir in os.listdir():
    if(dir.startswith("jdk")):
      is_java_installed = True

  # # If not there, install it
  if(not is_java_installed):
    jdk.install('11', path="./")

  # print(is_java_installed)
  #Set location of Java installation
  os.environ["JAVA_HOME"] = "./"

#####################################################
#      End Java Install / Local Testing
#####################################################

  # Start Java Virtual Machine
  # Needed in Container
  jpype.startJVM()


@app.route("/", methods=["GET","POST"])
def send_excel():
  if request.method == "GET":
    return "TESTING HTTPS"
  else:

#  Parameters determine format of Excel sheet
    excel_data = request.get_json().get('data')

    # g module allows variables to accessed globally within the same response
    # Doesn't function like a normal Global
    g.title = request.get_json().get('excelFormattingParameters')[0]['fileTitle']
    highlight_rows = request.get_json().get('excelFormattingParameters')[0]['highlight']
    column_widths = request.get_json().get('excelFormattingParameters')[0]['widths']
    excel_sheet_title = request.get_json().get('excelFormattingParameters')[0]['excelTitle']
    date_range = request.get_json().get('excelFormattingParameters')[0]['dateRange']
    pdf_or_no = request.get_json().get('excelFormattingParameters')[0]['pdf']

    output = make_excel_sheet(g.get('title'), excel_data, highlight_rows, column_widths, excel_sheet_title, date_range, pdf_or_no)


  # If the user wants a PDF the make_excel_sheet function will return 0 as a value
  if(output == 0):
    return_data = convert_to_pdf(g.get("title"))

  #Sends the file as a response to the browser to be downloaded
  #download_name has to be set even though it won't be used, the browser will normally handle the name of the file
  return send_file(output if output != 0 else return_data, download_name="not_needed.pdf" if output == 0 else "not_needed.xlsx", as_attachment=True)


@app.after_request
def after_request_func(response):

    """ 
    This function will run after a request, as long as no exceptions occur.
    It must take and return the same parameter - an instance of response_class.
    This is a good place to do some application cleanup.
    """

    close = Response("Testing")

    # Will perform some clean up of local files
    def process_after_request():

      if os.path.isfile("xlsx-to-pdf.pdf"):
        os.remove("xlsx-to-pdf.pdf")
      
      if os.path.isfile("xlsx-to-pdf-cropped.pdf"):
        os.remove("xlsx-to-pdf-cropped.pdf")
      
      if(g.get("title")):
        if os.path.isfile(g.get("title") + ".xlsx"):
          os.remove(g.get("title") + ".xlsx")

    # Will call function at end of request
    close.call_on_close(process_after_request())

    return response


app.config.update(
  HEALTHZ = {
    "live": "app.liveness",
    "ready": "app.readiness",
  }
)
  

if __name__ == '__main__':
  app.run(host="0.0.0.0", ssl_context='adhoc')