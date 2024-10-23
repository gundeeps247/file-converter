import os
from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash
from werkzeug.utils import secure_filename
from pdf2docx import Converter
import subprocess

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['CONVERTED_FOLDER'] = 'converted_files/'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Max file size: 16MB
app.secret_key = 'supersecretkey'

# Allowed file extensions
ALLOWED_EXTENSIONS = {'pdf', 'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)

    file = request.files['file']

    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        if filename.endswith('.pdf'):
            converted_file = pdf_to_word(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        elif filename.endswith('.docx'):
            converted_file = word_to_pdf(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        else:
            flash('Unsupported file format')
            return redirect(url_for('index'))
        
        return redirect(url_for('download_file', filename=converted_file))

    else:
        flash('Invalid file format')
        return redirect(url_for('index'))

@app.route('/uploads/<filename>')
def download_file(filename):
    return send_from_directory(app.config['CONVERTED_FOLDER'], filename, as_attachment=True)

def pdf_to_word(pdf_path):
    word_filename = os.path.splitext(os.path.basename(pdf_path))[0] + '.docx'
    word_path = os.path.join(app.config['CONVERTED_FOLDER'], word_filename)

    # Convert PDF to Word using pdf2docx to preserve layout
    cv = Converter(pdf_path)
    cv.convert(word_path, start=0, end=None)  # Convert entire document
    cv.close()

    return word_filename

def word_to_pdf(docx_path):
    pdf_filename = os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
    pdf_path = os.path.join(app.config['CONVERTED_FOLDER'], pdf_filename)

    # Use LibreOffice's command-line interface to convert DOCX to PDF
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', app.config['CONVERTED_FOLDER'], docx_path])

    return pdf_filename

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['CONVERTED_FOLDER'], exist_ok=True)
    app.run(debug=True)
