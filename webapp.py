from flask import Flask, request, render_template_string, send_file
from fillpdf import fillpdfs
import pandas as pd
import os

app = Flask(__name__)

# Ensure directories exist
os.makedirs('Pdfs', exist_ok=True)
os.makedirs('uploads', exist_ok=True)
os.makedirs('Export', exist_ok=True)

# Paths to the PDF templates
pdf_paths = {
    "HKT_Service_Application_Form": "Pdfs/HKT Payment Merchant Service Application Form.pdf",
    "HKT_POS_Moblie_Form": "Pdfs/SmartPOS_SoftPOS_Application_Agreement.pdf"
}

# HTML template for file upload with improved styling
upload_form = '''
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>HKT PDF Filler</title>
    <style>
      body {
        font-family: Arial, sans-serif;
        background-color: #f4f4f4;
        display: flex;
        justify-content: center;
        align-items: center;
        height: 100vh;
        margin: 0;
      }
      .container {
        background-color: #fff;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        text-align: center;
      }
      h1 {
        color: #333;
      }
      input[type="file"] {
        margin: 20px 0;
      }
      input[type="submit"] {
        background-color: #0052bd;
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 4px;
        cursor: pointer;
      }
      input[type="submit"]:hover {
        background-color: #7cb4fc;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <h1>Upload Excel File</h1>
      <form action="/" method="post" enctype="multipart/form-data">
        <input type="file" name="file" accept=".xlsx">
        <select name="form_type">
          <option value="HKT_Service_Application_Form">HKT Service Application Form</option>
          <option value="HKT_POS_Moblie_Form">HKT POS Mobile Form</option>
        </select>
        <input type="submit" value="Upload">
      </form>
    </div>
  </body>
</html>
'''

def read_excel_data(excel_path, sheet_name):
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    data = {}
    for index, row in df.iterrows():
        field_name = str(row['Field Name'])  # Assuming field names are in column 'Field Name'
        field_value = row['Input']  # Assuming input values are in column 'Input'
        if pd.isna(field_value):  # Check if the value is NaN
            field_value = ''  # Replace NaN with an empty string
        data[field_name] = field_value
    return data

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        form_type = request.form['form_type']
        if file and form_type in pdf_paths:
            # Save the uploaded file
            excel_path = os.path.join('uploads', file.filename)
            file.save(excel_path)
            
            # Read the Excel file into a DataFrame
            df = pd.read_excel(excel_path, sheet_name=form_type)
            
            # Process the DataFrame
            data = {}
            for index, row in df.iterrows():
                field_name = str(row['Field Name'])  # Assuming field names are in column 'Field Name'
                field_value = row['Input']  # Assuming input values are in column 'Input'
                if pd.isna(field_value):  # Check if the value is NaN
                    field_value = ''  # Replace NaN with an empty string
                data[field_name] = field_value
            
            # Fill PDF fields with data
            pdf_path = pdf_paths[form_type]
            output_path = os.path.join('Export', f"Filled_{os.path.basename(pdf_path)}")
            fillpdfs.write_fillable_pdf(pdf_path, output_path, data)
            
            # Provide the filled PDF for download
            return send_file(output_path, as_attachment=True)
    
    return render_template_string(upload_form)

if __name__ == '__main__':
    app.run(debug=True)