from flask import Flask, render_template, request, send_file, redirect, url_for
from flask_bootstrap import Bootstrap
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileRequired, FileAllowed
from wtforms import SubmitField
from io import BytesIO
import pandas as pd
import zipfile
import os
import subprocess
from feedback_analysis import analyze_feedback, generate_charts, generate_markdown_report, export_to_excel
import pdfkit

def generate_pdf(markdown_file):
    pdf_filename = 'feedback_report.pdf'
    subprocess.run(['pandoc', markdown_file, '-o', pdf_filename, '--pdf-engine=wkhtmltopdf', '--pdf-engine-opt=--enable-local-file-access', '--css=github.css'])
    return pdf_filename

def generate_pdf_latex(markdown_file):
    pdf_filename = 'feedback_report_latex.pdf'
    # subprocess.run(['pandoc', markdown_file, '-o', pdf_filename, '--pdf-engine=xelatex', '-N', '--shift-heading-level-by=1'])
    subprocess.run(['pandoc', markdown_file, '-o', pdf_filename, '--pdf-engine=xelatex', '-N'])
    return pdf_filename

def generate_zip(markdown_file, charts_dir):
    zip_filename = 'feedback_report.zip'
    with zipfile.ZipFile(zip_filename, 'w') as zip_file:
        zip_file.write(markdown_file)
        for root, dirs, files in os.walk(charts_dir):
            for file in files:
                zip_file.write(os.path.join(root, file))
    return zip_filename

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key'
Bootstrap(app)


class UploadForm(FlaskForm):
    file = FileField('Upload CSV', validators=[FileRequired(), FileAllowed(['csv'], 'CSV files only!')])
    submit = SubmitField('Upload')

@app.route('/', methods=['GET', 'POST'])
def index():
    form = UploadForm()
    if form.validate_on_submit():
        file = form.file.data
        file_content = file.read().decode('utf-8')
        analysis_result = analyze_feedback(file_content)
        charts_dir = './static/charts'
        generate_charts(analysis_result, charts_dir)
        markdown_report = 'feedback_report.md'
        generate_markdown_report(analysis_result, markdown_report, charts_dir)
        generate_zip(markdown_report, charts_dir)
        excel_report = 'feedback_report.xlsx'
        export_to_excel(analysis_result, excel_report, file_content)
        generate_pdf(markdown_report)
        generate_pdf_latex(markdown_report)
        return redirect(url_for('report'))
    return render_template('index.html', form=form)


@app.route('/download_sample')
def download_sample():
    sample_data = '''Year,Odd_Even,Branch,Sem,Responce_Count,Term_Start,Term_End,Subject_Code,Subject_ShortForm,Subject_FullName,Faculty_Initial,Faculty_Name,Q1,Q2,Q3,Q4,Q5,Q6,Q7,Q8,Q9,Q10,Q11,Q12
2023,Odd,EC,5,5,27/07/23,16/12/23,4300021,E&S,Entrepreneurship and Start-ups,SPJ,Mr. S P Joshiara,5,5,5,5,5,5,5,5,5,5,5,5
'''
    sample_csv = BytesIO(sample_data.encode('utf-8'))
    return send_file(sample_csv, download_name='sample_feedback.csv', as_attachment=True)

@app.route('/report')
def report():
    return render_template('report.html')

@app.route('/download_markdown')
def download_markdown():
    return send_file('feedback_report.zip', as_attachment=True)

@app.route('/download_excel')
def download_excel():
    return send_file('feedback_report.xlsx', as_attachment=True)

@app.route('/download_pdf')
def download_pdf():
    return send_file('feedback_report.pdf', as_attachment=True)

@app.route('/download_pdf_latex')
def download_pdf_latex():
    return send_file('feedback_report_latex.pdf', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)