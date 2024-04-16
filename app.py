from flask import Flask, flash, request, render_template, redirect, send_file
from werkzeug.utils import secure_filename
import pandas as pd
import re
from docx import Document
from PyPDF2 import PdfReader

ALLOWED_EXTENSIONS = {'docx', 'pdf'}

app = Flask(__name__)

app.secret_key = '761215@Om'

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_docx(file):
    doc = Document(file)
    
    text_content = ""
    for paragraph in doc.paragraphs:
        text_content += paragraph.text + "\n"
        
    return text_content

def extract_text_from_pdf(file):
    pdf = PdfReader(file)
        
    text_content = ""
    for page in pdf.pages:
        text_content += page.extract_text()
            
    return text_content

def extract_details(text):
    contact_match = '\d{3}-\d{3}-\d{4}|\d{10}'
    email_match = '[a-z0-9A-Z._]*@[a-z0-9A-Z._]*\.[a-zA-Z]*'

    contacts = re.findall(contact_match, text)
    emails = re.findall(email_match, text)
    
    if len(contacts)>0:
        contact = contacts[0]
    else:
        contact = None
    if len(emails)>0:
        email = emails[0]
    else:
        email = None
    
    return [contact, email]

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        uploaded_files = request.files.getlist('files')
        df = pd.DataFrame()

        for file in uploaded_files:
            if file.filename == '':
                flash('No selected file')
                return redirect(request.url)
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                if filename.rsplit('.', 1)[1].lower() == 'docx':
                    doc_text = extract_text_from_docx(file)
                elif filename.rsplit('.', 1)[1].lower() == 'pdf':
                    doc_text = extract_text_from_pdf(file)

                output = extract_details(doc_text)
                output.append(doc_text)
                data = pd.DataFrame([output], columns = ['Contacts', 'Emails', 'Text'])
                df = pd.concat([df,data], ignore_index = True)
                
        df.to_excel("Excel.xlsx", index = False)

    return render_template('index.html')
    
@app.route('/download', methods=['GET'])
def download_file():
    file_path = "Excel.xlsx"
    return send_file(file_path, as_attachment=True)
    
if __name__ == "__main__":
    app.run(debug=True)