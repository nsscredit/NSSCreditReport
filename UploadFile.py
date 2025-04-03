
from flask import Flask, request, redirect, url_for, send_file, Response
import os
import re
import glob
import subprocess

app = Flask(__name__)
#UPLOAD_FOLDER = os.getcwd()
#os.makedirs(UPLOAD_FOLDER, exist_ok=True)
#UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")

UPLOAD_FOLDER = "/tmp/uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)  # Ensure the folder exists




@app.route('/')
def upload_form():
    message = request.args.get('message', '')
    message_class = request.args.get('message_class', '')
    return f""" 

    <!doctype html>
    <html>
    <head>
      <title>Volunteer Credit Report</title>
      <style>
        body {{ font-family: Arial, sans-serif; background-color: #f4f4f4; text-align: center; padding: 20px; }}
        h1 {{ color: #333; }}
        form {{ margin: 20px auto; padding: 20px; width: 300px; background: #fff; box-shadow: 0 0 10px rgba(0,0,0,0.1); border-radius: 10px; }}
        input[type="file"] {{ margin-bottom: 10px; }}
        input[type="submit"], button {{ background-color: #4CAF50; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; }}
        input[type="submit"]:hover, button:hover {{ background-color: #45a049; }}
        #loading {{ display: none; color: #333; margin-top: 10px; }}
        #message {{ margin-top: 10px; font-weight: bold; padding: 10px; border-radius: 5px; }}
        .success {{ color: #155724; background-color: #d4edda; border: 1px solid #c3e6cb; }}
        .error {{ color: #721c24; background-color: #f8d7da; border: 1px solid #f5c6cb; }}
      </style>
       <script>
            function displayFileNames() {{
                var input = document.getElementById('fileInput');
                var fileNamesDiv = document.getElementById('fileNames');
                
                if (input.files.length > 0) {{
                    var fileList = "<strong>Selected files:</strong><br>";
                    for (var i = 0; i < input.files.length; i++) {{
                        fileList += input.files[i].name + "<br>";
                    }}
                    fileList += "</ul>";
                    fileNamesDiv.innerHTML = fileList;
                }} else {{
                    fileNamesDiv.innerHTML = "";
                }}
            }}
      <script>
        function showLoading() {{
          document.getElementById('loading').style.display = 'block';
        }}
      </script>
    </head>
    
    <body>
    <h1>NSS Credit Report</h1>
    <h2>Upload one or more Credit Excel data to generate Credit Report</h2>
    
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="files" multiple id="fileInput" onchange="displayFileNames()">
        <br><br>
        <div id="fileNames"></div> <!-- Area to display selected file names -->
        <br>
        <input type="submit" value="Upload">
    </form>
    
    <br>
    <form action="/process" method="post" onsubmit="showLoading()">
        <button type="submit">Generate Report</button>
    </form>
    
    <div id="loading" style="display: none;">Processing... Please wait!</div>
    {'<div id="message" class="' + message_class + '">' + message + '</div>' if message else ''}

        </body>
    </html>
    """



@app.route('/upload', methods=['POST'])
def upload_file():
    if 'files' not in request.files:
        return redirect(url_for('upload_form', message='No files found!', message_class='error'))
    files = request.files.getlist('files')

    if not files or all(file.filename == '' for file in files):
        return redirect(url_for('upload_form', message='No files uploaded!', message_class='error'))
    uploaded_files = []  # List to store uploaded file names
    for file in files:
        if file.filename.endswith(('.xls', '.xlsx')):
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)
            uploaded_files.append(file.filename)
	
    # Convert list of file names into a string message
    print(f"Saving file to: {UPLOAD_FOLDER}")  # DEBUG: Print path

    message = f"Files uploaded successfully: {', '.join(uploaded_files)}"

    return redirect(url_for('upload_form', message=message, message_class='success'))
    # for file in files:
    #    if file.filename.endswith(('.xls', '.xlsx')):
    #        file.save(os.path.join(UPLOAD_FOLDER, file.filename))

    #return redirect(url_for('upload_form', message='Files uploaded successfully!', message_class='success'))


@app.route('/process', methods=['POST'])
def process_files():
    try:
        subprocess.run(['python', 'Report.py'], check=True)
        return redirect(url_for('download_report'))
    except subprocess.CalledProcessError:
        return 'Error running Report.py'


@app.route('/download_report')
def download_report():
    report_path = os.path.join(os.getcwd(), 'credit_report.pdf')
    if os.path.exists(report_path):
        #return send_file(report_path, as_attachment=False, download_name="Credit Report.pdf")
        response = Response(open(report_path, 'rb'), mimetype='application/pdf')
        response.headers['Content-Disposition'] = 'inline; filename="Credit Report.pdf"'
        response.headers['Content-Type'] = 'application/pdf'
        return response
    else:
        return 'Report not found!'


if __name__ == '__main__':
    app.run(debug=True)
