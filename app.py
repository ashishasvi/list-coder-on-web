# app.py
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from file_processor import process_files  # We'll define process_files in file_processor.py

app = Flask(__name__)

app.secret_key = "supersecretkey"  # required for flash messages

@app.route('/')
def index():
    # Renders a simple upload form with two <input type="file"> fields
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    # Ensure user has posted files
    if 'csvfile' not in request.files or 'excelfile' not in request.files:
        flash("Please upload both a CSV (Row File) and an Excel (Lufthansa File).")
        return redirect(url_for('index'))

    row_file = request.files['csvfile']       # This is the Row CSV
    lufthansa_file = request.files['excelfile']  # This is the Lufthansa XLSX
    condition_codes=request.form.get("text_values")

    if not row_file or not lufthansa_file:
        flash("Please upload both a CSV (Row File) and an Excel (Lufthansa File).")
        return redirect(url_for('index'))

    # Process the two files in memory (mimicking your VBA pipeline)
    final_excel = process_files(row_file, lufthansa_file,condition_codes)

    # Return the processed Excel as a download
    return send_file(
        final_excel,
        as_attachment=True,
        download_name=lufthansa_file.filename,  # or whichever filename you wish
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(port=5000)
