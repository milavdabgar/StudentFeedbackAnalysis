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
    data = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(list)))))
    
    csv_data = StringIO(file_content)
    csv_reader = csv.DictReader(csv_data)
    
    for row in csv_reader:
        year_term = f"{row['Odd_Even']} - {row['Year']}"
        branch = row['Branch']
        semester = row['Sem']
        subject = f"{row['Subject_ShortForm']} ({row['Subject_Code']})"
        faculty = row['Faculty_Name']
        
        scores = [int(row[f'Q{i}']) for i in range(1, 13)]
        data[year_term][branch][semester][subject][faculty].append(scores)
    
    analysis = {}
    
    # Year-Term Analysis
    analysis['Year-Term Analysis'] = {
        year_term: {
            f"{branch} - Sem {semester}": calculate_average([s for sub_scores in scores for s in sub_scores]) 
            for branch in data[year_term] for semester in data[year_term][branch] 
            for subject in data[year_term][branch][semester] for faculty in data[year_term][branch][semester][subject]
            for scores in [data[year_term][branch][semester][subject][faculty]]
        } for year_term in data
    }
    
    # Branch Analysis
    analysis['Branch Analysis'] = {
        branch: {
            'Average scores for Q1-Q12': [calculate_average([scores[i] for scores in branch_scores]) for i in range(12)],
            'Overall average': calculate_average([s for sub_scores in branch_scores for s in sub_scores])
        } for year_term in data for branch in data[year_term] 
        for branch_scores in [[score for semester in data[year_term][branch] for subject in data[year_term][branch][semester]
                               for faculty in data[year_term][branch][semester][subject] for score in data[year_term][branch][semester][subject][faculty]]]
    }
    
    # Subject Analysis
    analysis['Subject Analysis'] = {
        subject: {
            'Overall average': calculate_average([s for sub_scores in scores for s in sub_scores])
        } for year_term in data for branch in data[year_term] for semester in data[year_term][branch] 
        for subject in data[year_term][branch][semester]
        for scores in [[score for faculty in data[year_term][branch][semester][subject] for score in data[year_term][branch][semester][subject][faculty]]]
    }
    
    # Faculty Analysis
    analysis['Faculty Analysis'] = {}
    for year_term in data:
        for branch in data[year_term]:
            for semester in data[year_term][branch]:
                for subject in data[year_term][branch][semester]:
                    for faculty in data[year_term][branch][semester][subject]:
                        scores = data[year_term][branch][semester][subject][faculty]
                        if faculty not in analysis['Faculty Analysis']:
                            analysis['Faculty Analysis'][faculty] = {'Subjects': {}, 'Overall average': []}
                        analysis['Faculty Analysis'][faculty]['Subjects'][subject] = calculate_average([s for sub_scores in scores for s in sub_scores])
                        analysis['Faculty Analysis'][faculty]['Overall average'].extend([s for sub_scores in scores for s in sub_scores])
    
    for faculty in analysis['Faculty Analysis']:
        analysis['Faculty Analysis'][faculty]['Overall average'] = calculate_average(analysis['Faculty Analysis'][faculty]['Overall average'])
    
    # Parameter-wise Analysis
    analysis['Parameter-wise Analysis'] = {}
    
    for category in ['Branch', 'Subject', 'Faculty']:
        analysis['Parameter-wise Analysis'][category] = {}
        for year_term in data:
            for branch in data[year_term]:
                for semester in data[year_term][branch]:
                    for subject in data[year_term][branch][semester]:
                        for faculty in data[year_term][branch][semester][subject]:
                            scores = data[year_term][branch][semester][subject][faculty]
                            key = branch if category == 'Branch' else subject if category == 'Subject' else faculty
                            if key not in analysis['Parameter-wise Analysis'][category]:
                                analysis['Parameter-wise Analysis'][category][key] = {f'Q{i}': [] for i in range(1, 13)}
                            for i in range(12):
                                analysis['Parameter-wise Analysis'][category][key][f'Q{i+1}'].extend([s[i] for s in scores])
        
        for key in analysis['Parameter-wise Analysis'][category]:
            for param in analysis['Parameter-wise Analysis'][category][key]:
                analysis['Parameter-wise Analysis'][category][key][param] = calculate_average(analysis['Parameter-wise Analysis'][category][key][param])
    
    return analysis

def generate_charts(analysis_result, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for category in ['Branch', 'Subject', 'Faculty']:
        keys = list(analysis_result[f'{category} Analysis'].keys())
        overall_averages = [data['Overall average'] for data in analysis_result[f'{category} Analysis'].values()]
        plt.figure(figsize=(12, 6))
        plt.bar(keys, overall_averages)
        plt.xlabel(category)
        plt.ylabel('Overall Average')
        plt.title(f'{category} Analysis')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        plt.savefig(f'{output_dir}/{category.lower()}_analysis.png')
        plt.close()     

def generate_excel_report(analysis_result, output_file, original_data):
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    
    original_df = pd.read_csv(StringIO(original_data))
    original_df.to_excel(writer, sheet_name='Original Data', index=False)
    
    for sheet_name, data in analysis_result.items():
        if sheet_name == 'Year-Term Analysis':
            df = pd.DataFrame.from_dict({k: {k2: v2 for k2, v2 in v.items()} for k, v in data.items()}, orient='index')
        elif sheet_name in ['Branch Analysis', 'Subject Analysis', 'Faculty Analysis']:
            df = pd.DataFrame.from_dict(data, orient='index')
        elif sheet_name == 'Parameter-wise Analysis':
            for category, category_data in data.items():
                df = pd.DataFrame.from_dict(category_data, orient='index')
                df.to_excel(writer, sheet_name=f'{category} Parameter-wise')
        
        if sheet_name != 'Parameter-wise Analysis':
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
        
        file.write('### Year-Term Analysis\n\n')
        for year_term, data in analysis_result['Year-Term Analysis'].items():
            file.write(f'### {year_term}\n\n')
            file.write('| Branch - Semester | Average Score |\n')
            file.write('|-------------------|---------------|\n')
            for key, value in data.items():
                file.write(f'| {key} | {value:.2f} |\n')
            file.write('\n')
        
        file.write('### Branch Analysis\n\n')
        file.write('| Branch | Overall Average | Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |\n')
        file.write('|--------|-----------------|----|----|----|----|----|----|----|----|----|----|-----|-----|\n')
        for branch, data in analysis_result['Branch Analysis'].items():
            file.write(f'| {branch} | {data["Overall average"]:.2f} |')
            for score in data['Average scores for Q1-Q12']:
                file.write(f' {score:.2f} |')
            file.write('\n')
        file.write('\n')
        file.write('![Branch Analysis](static/images/charts/branch_analysis.png)\n\n')
        
        file.write('### Subject Analysis\n\n')
        file.write('| Subject | Overall Average |\n')
        file.write('|---------|------------------|\n')
        for subject, data in analysis_result['Subject Analysis'].items():
            file.write(f'| {subject} | {data["Overall average"]:.2f} |\n')
        file.write('\n')
        file.write('![Subject Analysis](static/images/charts/subject_analysis.png)\n\n')
        
        file.write('### Faculty Analysis\n\n')
        for faculty, data in analysis_result['Faculty Analysis'].items():
            file.write(f'#### {faculty}\n\n')
            file.write(f'- Overall Average: {data["Overall average"]:.2f}\n\n')
            file.write('| Subject | Average Score |\n')
            file.write('|---------|---------------|\n')
            for subject, average in data['Subjects'].items():
                file.write(f'| {subject} | {average:.2f} |\n')
            file.write('\n')
        file.write('![Faculty Analysis](static/images/charts/faculty_analysis.png)\n\n')
        
        file.write('### Parameter-wise Analysis\n\n')
        for category in ['Branch', 'Subject', 'Faculty']:
            file.write(f'#### Parameter-wise {category} Analysis\n\n')
            file.write(f'| {category} | Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |\n')
            file.write('|----------|----|----|----|----|----|----|----|----|----|-----|-----|-----|\n')
            for key, data in analysis_result['Parameter-wise Analysis'][category].items():
                file.write(f'| {key} |')
                for param in [f'Q{i}' for i in range(1, 13)]:
                    file.write(f' {data[param]:.2f} |')
                file.write('\n')
            file.write('\n')
    
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
    generate_charts(analysis_result, output_dir)
    
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