import os
import io
import uuid
import win32com.client
import pythoncom
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

def convertExcelToPdf(excel_bytes):
    # Initialize COM
    pythoncom.CoInitialize()

    # Generate unique IDs for temporary files
    unique_id = str(uuid.uuid4())

    # Create a temporary file for the Excel file
    with io.BytesIO(excel_bytes) as excel_file:
        temp_excel_path = f'temp_excel_{unique_id}.xlsx'
        with open(temp_excel_path, 'wb') as f:
            f.write(excel_file.getbuffer())

        # Initialize Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        # Open the Excel file
        wb = excel.Workbooks.Open(os.path.abspath(temp_excel_path))

        # Define the temporary PDF file path
        temp_pdf_path = f'temp_pdf_{unique_id}.pdf'

        # Export as PDF
        wb.ExportAsFixedFormat(0, os.path.abspath(temp_pdf_path))

        # Close the Excel workbook
        wb.Close(False)
        excel.Quit()

    # Read the generated PDF file and return as byte array
    with open(temp_pdf_path, 'rb') as f:
        pdf_bytes = f.read()

    # Clean up temporary files
    os.remove(temp_excel_path)
    os.remove(temp_pdf_path)

    return pdf_bytes

@app.route('/convert', methods=['POST'])
def convert():
    # Get the Excel file from the request
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    # Read the Excel file as byte array
    excel_bytes = file.read()

    # Convert to PDF byte array
    pdf_bytes = convertExcelToPdf(excel_bytes)

    # Return the PDF as a byte array
    return send_file(io.BytesIO(pdf_bytes), download_name='output.pdf', as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080,debug=True)
