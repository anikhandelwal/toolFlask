import os
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import fitz
from docx import Document
from docx.shared import Pt

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def pdf_to_word(pdf_path):
    pdf_document = fitz.open(pdf_path)
    word_document = Document()

    for page_number in range(pdf_document.page_count):
        page = pdf_document[page_number]
        for block in page.get_text("blocks"):
            text = block[4]
            font_size = Pt(12)
            is_bold = False
            paragraph = word_document.add_paragraph(text)

            for run in paragraph.runs:
                run.font.size = font_size

            if is_bold:
                paragraph.runs[0].bold = True

        if page_number < pdf_document.page_count - 1:
            word_document.add_page_break()

    # Get the "Downloads" directory path
    downloads_dir = os.path.expanduser('~/Downloads')
    
    # Use the same filename as the input PDF in the "Downloads" directory
    word_path = os.path.join(downloads_dir, os.path.basename(pdf_path).rsplit('.', 1)[0] + '.docx')

    word_document.save(word_path)
    pdf_document.close()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('index.html', error='No file part')

        file = request.files['file']

        if file.filename == '':
            return render_template('index.html', error='No selected file')

        if file and allowed_file(file.filename):
            # Ensure the 'uploads' directory exists
            if not os.path.exists(app.config['UPLOAD_FOLDER']):
                os.makedirs(app.config['UPLOAD_FOLDER'])

            file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(file_path)

            pdf_to_word(file_path)

            # Use the same filename as the input PDF in the "Downloads" directory
            output_path = os.path.join(os.path.expanduser('~/Downloads'), os.path.basename(file_path).rsplit('.', 1)[0] + '.docx')

            return render_template('index.html', success=True, output_path=output_path)

    return render_template('index.html')

@app.route('/download')
def download():
    output_path = request.args.get('output_path', '')
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=False)
