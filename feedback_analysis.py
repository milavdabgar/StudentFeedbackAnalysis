import csv
from collections import defaultdict
import matplotlib.pyplot as plt
import pandas as pd
from io import StringIO 
import zipfile
import subprocess
import os

def calculate_average(scores):
    return sum(scores) / len(scores)

def analyze_feedback(file_content):
    data = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(list))))
    
    csv_data = StringIO(file_content)
    csv_reader = csv.DictReader(csv_data)
    
    for row in csv_reader:
        year_term = f"{row['Odd_Even']} - {row['Year']}"
        branch = row['Branch']
        semester = f"{branch} - Sem {row['Sem']}"
        subject = f"{row['Subject_ShortForm']} ({row['Subject_Code']})"
        faculty = row['Faculty_Name']
        
        scores = [int(row[f'Q{i}']) for i in range(1, 13)]
        data[year_term][branch][semester][subject].append((faculty, scores))
    
    analysis = {}
    
    # Faculty Analysis
    analysis['Faculty Analysis'] = {}
    for year_term in data:
        for branch in data[year_term]:
            for semester in data[year_term][branch]:
                for subject, faculty_scores in data[year_term][branch][semester].items():
                    for faculty, scores in faculty_scores:
                        if faculty not in analysis['Faculty Analysis']:
                            analysis['Faculty Analysis'][faculty] = {'Subjects': {}, 'Overall average': 0}
                        analysis['Faculty Analysis'][faculty]['Subjects'][subject] = calculate_average(scores)
    
    for faculty in analysis['Faculty Analysis']:
        subject_averages = list(analysis['Faculty Analysis'][faculty]['Subjects'].values())
        analysis['Faculty Analysis'][faculty]['Overall average'] = calculate_average(subject_averages)
    
    # Subject Analysis
    analysis['Subject Analysis'] = {}
    for year_term in data:
        for branch in data[year_term]:
            for semester in data[year_term][branch]:
                for subject, faculty_scores in data[year_term][branch][semester].items():
                    if subject not in analysis['Subject Analysis']:
                        analysis['Subject Analysis'][subject] = {'Faculties': {}, 'Overall average': 0}
                    for faculty, scores in faculty_scores:
                        analysis['Subject Analysis'][subject]['Faculties'][faculty] = calculate_average(scores)
    
    for subject in analysis['Subject Analysis']:
        faculty_averages = list(analysis['Subject Analysis'][subject]['Faculties'].values())
        analysis['Subject Analysis'][subject]['Overall average'] = calculate_average(faculty_averages)
    
    # Semester Analysis
    analysis['Semester Analysis'] = {}
    for year_term in data:
        for branch in data[year_term]:
            for semester in data[year_term][branch]:
                if semester not in analysis['Semester Analysis']:
                    analysis['Semester Analysis'][semester] = []
                for subject, faculty_scores in data[year_term][branch][semester].items():
                    for _, scores in faculty_scores:
                        analysis['Semester Analysis'][semester].extend(scores)
    
    for semester in analysis['Semester Analysis']:
        analysis['Semester Analysis'][semester] = calculate_average(analysis['Semester Analysis'][semester])
    
    # Branch Analysis
    analysis['Branch Analysis'] = {}
    for year_term in data:
        for branch in data[year_term]:
            if branch not in analysis['Branch Analysis']:
                analysis['Branch Analysis'][branch] = []
            for semester in data[year_term][branch]:
                for subject, faculty_scores in data[year_term][branch][semester].items():
                    for _, scores in faculty_scores:
                        analysis['Branch Analysis'][branch].extend(scores)
    
    for branch in analysis['Branch Analysis']:
        analysis['Branch Analysis'][branch] = calculate_average(analysis['Branch Analysis'][branch])
    
    return analysis
           
def generate_charts(analysis_result, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for category in ['Branch', 'Subject', 'Faculty']:
        if category == 'Faculty':
            keys = list(analysis_result[f'{category} Analysis'].keys())
            overall_averages = [data['Overall average'] for data in analysis_result[f'{category} Analysis'].values()]
        else:
            keys = list(analysis_result[f'{category} Analysis'].keys())
            overall_averages = list(analysis_result[f'{category} Analysis'].values())

        # Convert keys to strings
        keys = [str(key) for key in keys]

        fig, ax = plt.subplots(figsize=(12, 6))
        ax.bar(keys, overall_averages)
        ax.set_xlabel(category)
        ax.set_ylabel('Overall Average')
        ax.set_title(f'{category} Analysis')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        plt.savefig(f'{output_dir}/{category.lower()}_analysis.png')
        plt.close(fig)

def generate_excel_report(analysis_result, output_file, original_data):
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    
    original_df = pd.read_csv(StringIO(original_data))
    original_df.to_excel(writer, sheet_name='Original Data', index=False)
    
    for sheet_name, data in analysis_result.items():
        if sheet_name == 'Faculty Analysis':
            df = pd.DataFrame.from_dict({faculty: {subject: score for subject, score in faculty_data['Subjects'].items()} for faculty, faculty_data in data.items()}, orient='index')
            df['Overall Average'] = [faculty_data['Overall average'] for faculty_data in data.values()]
        elif sheet_name == 'Subject Analysis':
            df = pd.DataFrame.from_dict({subject: {faculty: score for faculty, score in subject_data['Faculties'].items()} for subject, subject_data in data.items()}, orient='index')
            df['Overall Average'] = [subject_data['Overall average'] for subject_data in data.values()]
        elif sheet_name == 'Semester Analysis':
            df = pd.Series(data).to_frame(name='Average Score')
        elif sheet_name == 'Branch Analysis':
            df = pd.Series(data).to_frame(name='Average Score')
        
        df.to_excel(writer, sheet_name=sheet_name)
    
    writer._save()

def generate_markdown_report(analysis_result, output_file):
    with open(output_file, 'w') as file:
        file.write("# Student Feedback Analysis Report\n\n")
        
        file.write("## Assessment Parameters & Rating Scale\n\n")
        file.write("### Assessment Parameters\n\n")
        file.write("- **Q1 Syllabus Coverage:** Has the Teacher covered entire Syllabus as prescribed by University/ College/ Board?\n")
        file.write("- **Q2 Topics Beyond Syllabus:** Has the Teacher covered relevant topics beyond syllabus?\n")
        file.write("- **Q3 Pace of Teaching:** Pace on which contents were covered?\n")
        file.write("- **Q4 Practical Demo:** Support for the development of Student's skill (Practical demonstration)\n")
        file.write("- **Q5 Hands-on Training:** Support for the development of Student's skill (Hands-on training)\n")
        file.write("- **Q6 Technical Skills of Teacher:** Effectiveness of Teacher in terms of: Technical Skills\n")
        file.write("- **Q7 Communication Skills of Teacher:** Effectiveness of Teacher in terms of: Communication Skills\n")
        file.write("- **Q8 Doubt Clarification:** Clarity of expectations of students\n")
        file.write("- **Q9 Use of Teaching Tools:** Effectiveness of Teacher in terms of: Use of teaching aids\n")
        file.write("- **Q10 Motivation:** Motivation and inspiration for students to learn\n")
        file.write("- **Q11 Helpfulness of Teacher:** Willingness to offer help and advice to students\n")
        file.write("- **Q12 Student Progress Feedback:** Feedback provided on Student's progress\n\n")

        file.write("### Rating Scale\n\n")        
        file.write("Rating | Description\n")
        file.write("-------|------------\n")
        file.write("1      | Very Poor\n")
        file.write("2      | Poor\n")
        file.write("3      | Average\n")
        file.write("4      | Good\n")
        file.write("5      | Very Good\n\n")        

        file.write("## Feedback Analysis\n\n")            
        
        file.write("## Faculty Analysis\n\n")
        for faculty, data in analysis_result['Faculty Analysis'].items():
            file.write(f"### {faculty}\n\n")
            file.write(f"- Overall Average: {data['Overall average']:.2f}\n\n")
            file.write("| Subject | Average Score |\n")
            file.write("|---------|---------------|\n")
            for subject, average in data['Subjects'].items():
                file.write(f"| {subject} | {average:.2f} |\n")
            file.write("\n")
        
        file.write("## Subject Analysis\n\n")
        for subject, data in analysis_result['Subject Analysis'].items():
            file.write(f"### {subject}\n\n")
            file.write(f"- Overall Average: {data['Overall average']:.2f}\n\n")
            file.write("| Faculty | Average Score |\n")
            file.write("|---------|---------------|\n")
            for faculty, average in data['Faculties'].items():
                file.write(f"| {faculty} | {average:.2f} |\n")
            file.write("\n")
        
        file.write("## Semester Analysis\n\n")
        file.write("| Semester | Average Score |\n")
        file.write("|----------|---------------|\n")
        for semester, average in analysis_result['Semester Analysis'].items():
            file.write(f"| {semester} | {average:.2f} |\n")
        file.write("\n")
        
        file.write("## Branch Analysis\n\n")
        file.write("| Branch | Average Score |\n")
        file.write("|--------|---------------|\n")
        for branch, average in analysis_result['Branch Analysis'].items():
            file.write(f"| {branch} | {average:.2f} |\n")
        file.write("\n")
    
def generate_pdf_wkhtml(markdown_file):
    yaml_front_matter = '''---
title: Student Feedback Analysis Report
subtitle: EC Dept, Government Polytechnic Palanpur
margin-left: 2cm
margin-right: 2cm
margin-top: 2cm
margin-bottom: 2cm
---
'''
    with open('temp_report_wkhtml.md', 'w') as temp_file:
        temp_file.write(yaml_front_matter)
        with open(markdown_file, 'r') as original_file:
            temp_file.write(original_file.read())

    pdf_filename = 'feedback_report_wkhtml.pdf'
    subprocess.run(['pandoc', '-s', '-o', pdf_filename, '--pdf-engine=wkhtmltopdf', '--pdf-engine-opt=--enable-local-file-access', '--css=static/css/github.css', '--toc', '-N', '--shift-heading-level-by=-1', 'temp_report_wkhtml.md'])
    os.remove('temp_report_wkhtml.md')
    return pdf_filename


def generate_pdf_latex(markdown_file):
    yaml_front_matter = '''---
title: Student Feedback Analysis Report
subtitle: EC Dept, Government Polytechnic Palanpur
margin-left: 2.5cm
margin-right: 2.5cm
margin-top: 2cm
margin-bottom: 2cm
toc: True
# header-includes:
#   - |
#     ```{=latex}
#     \\usepackage{fontspec}
#     \\usepackage{polyglossia}
#     \\setmainlanguage{english}
#     \\setotherlanguage{sanskrit}
#     \\newfontfamily\\englishfont[Ligatures=TeX]{Noto Sans}
#     \\newfontfamily\\sanskritfont[Script=Gujarati]{Noto Sans Gujarati}
#     ```
---
'''
    with open('temp_report.md', 'w') as temp_file:
        temp_file.write(yaml_front_matter)
        with open(markdown_file, 'r') as original_file:
            temp_file.write(original_file.read())

    pdf_filename = 'feedback_report_latex.pdf'
    subprocess.run(['pandoc', '-s', '-o', pdf_filename, '--pdf-engine=xelatex', '-N', '--shift-heading-level-by=-1', 'temp_report.md'])
    os.remove('temp_report.md')
    return pdf_filename

def generate_report(analysis_result, original_data):
    output_dir = 'static/images/charts'
    # generate_charts(analysis_result, output_dir)
    
    markdown_file = 'feedback_report.md'
    generate_markdown_report(analysis_result, markdown_file)
    
    excel_file = 'feedback_report.xlsx'
    generate_excel_report(analysis_result, excel_file, original_data)
    
    pdf_wkhtml = generate_pdf_wkhtml(markdown_file)
    pdf_latex = generate_pdf_latex(markdown_file)
    
    zip_filename = 'feedback_report.zip'
    with zipfile.ZipFile(zip_filename, 'w') as zip_file:
        zip_file.write(markdown_file)
        zip_file.write(excel_file)
        zip_file.write(pdf_wkhtml)
        zip_file.write(pdf_latex)
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                zip_file.write(os.path.join(root, file))