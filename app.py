from flask import Flask, render_template, request, send_file, redirect, url_for
from flask_bootstrap import Bootstrap
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileRequired, FileAllowed
from wtforms import SubmitField
from io import BytesIO
from feedback_analysis import analyze_feedback, generate_report

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
        generate_report(analysis_result, file_content)
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

@app.route('/download_report')
def download_report():
    return send_file('feedback_report.zip', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)