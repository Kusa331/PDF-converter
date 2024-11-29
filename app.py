from flask import Flask, request, render_template, send_file
import os
from docx2pdf import convert

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'docx', 'doc'}


os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    if file and allowed_file(file.filename):
      
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)

    
        pdf_path = file_path.replace('.docx', '.pdf').replace('.doc', '.pdf')
        convert(file_path, pdf_path)

     
        return send_file(pdf_path, as_attachment=True)

    return "File type not allowed", 400

if __name__ == '__main__':
    app.run(debug=True)