import pandas as pd
from io import StringIO 
import zipfile
import subprocess
import os

def analyze_feedback(file_content):
    data = pd.read_csv(StringIO(file_content))
    data['Subject_Code'] = data['Subject_Code'].astype(str)  # Convert Subject_Code to string   

    subject_scores_faculty = data.groupby(['Subject_Code', 'Subject_ShortForm', 'Faculty_Name']).agg({
        **{f'Q{i}': 'mean' for i in range(1, 13)}
    }).reset_index()

    subject_scores_faculty['Average_Score'] = subject_scores_faculty[list(f'Q{i}' for i in range(1, 13))].mean(axis=1)

    subject_scores_overall = subject_scores_faculty.groupby(['Subject_Code', 'Subject_ShortForm'])['Average_Score'].mean().reset_index()
    subject_scores_overall.columns = ['Subject_Code', 'Subject_ShortForm', 'Overall_Average']

    faculty_scores_subject = subject_scores_faculty.groupby(['Faculty_Name', 'Subject_Code', 'Subject_ShortForm']).agg({
        **{f'Q{i}': 'mean' for i in range(1, 13)},
        'Average_Score': 'mean'
    }).reset_index()
    
    faculty_scores_overall = faculty_scores_subject.groupby('Faculty_Name')['Average_Score'].mean().reset_index()
    faculty_scores_overall.columns = ['Faculty_Name', 'Overall_Average']

    semester_scores = data.groupby(['Year', 'Term', 'Branch', 'Sem']).agg({
        **{f'Q{i}': 'mean' for i in range(1, 13)}
    }).reset_index()

    semester_scores['Branch_Semester'] = semester_scores['Branch'] + ' - ' + semester_scores['Sem'].astype(str)

    branch_scores = semester_scores.groupby('Branch').agg({
        **{f'Q{i}': 'mean' for i in range(1, 13)}
    })

    term_year_scores = semester_scores.groupby(['Year', 'Term']).agg({
        **{f'Q{i}': 'mean' for i in range(1, 13)}
    })

    correlation_matrix = faculty_scores_subject.pivot_table(index=['Subject_Code', 'Subject_ShortForm'], columns='Faculty_Name', values='Average_Score', aggfunc='mean')

    faculty_overall = faculty_scores_overall.set_index('Faculty_Name')['Overall_Average']
    correlation_matrix.loc['Faculty Overall'] = faculty_overall

    subject_overall = subject_scores_overall.set_index(['Subject_Code', 'Subject_ShortForm'])['Overall_Average']
    correlation_matrix['Subject Overall'] = subject_overall

    correlation_matrix = correlation_matrix.fillna('-')

    analysis_result = {
        'subject_scores_faculty': subject_scores_faculty,
        'subject_scores_overall': subject_scores_overall,
        'faculty_scores_subject': faculty_scores_subject,
        'faculty_scores_overall': faculty_scores_overall,
        'semester_scores': semester_scores,
        'branch_scores': branch_scores,
        'term_year_scores': term_year_scores,
        'correlation_matrix': correlation_matrix
    }

    return analysis_result

def generate_excel_report(analysis_result, output_file, original_data):
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    
    original_df = pd.read_csv(StringIO(original_data))
    original_df.to_excel(writer, sheet_name='Original Data', index=False)
    
    for sheet_name, data in analysis_result.items():
        df = pd.DataFrame(data)
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    writer._save()

def generate_markdown_report(analysis_result):
    def format_float(x):
        if isinstance(x, pd.Series):
            return x.apply(lambda val: f"{val:.2f}" if pd.notna(val) and isinstance(val, (int, float)) else val)
        else:
            return f"{x:.2f}" if pd.notna(x) and isinstance(x, (int, float)) else x

    report = "## Feedback Analysis\n\n"

    report += "### Branch Analysis (overall)\n\n"
    report += analysis_result['branch_scores'].mean(axis=1).reset_index().rename(columns={0: 'Average Score'}).apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "### Term-Year Analysis (overall)\n\n"
    report += analysis_result['term_year_scores'].mean(axis=1).reset_index().rename(columns={0: 'Average Score'}).apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "### Semester Analysis (overall)\n\n"
    report += analysis_result['semester_scores'].drop(columns=['Year', 'Term', 'Branch_Semester']).groupby(['Branch', 'Sem'])[list(f'Q{i}' for i in range(1, 13))].mean().mean(axis=1).reset_index().rename(columns={0: 'Score'})[['Branch', 'Sem', 'Score']].apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "### Subject Analysis (overall)\n\n"
    report += analysis_result['subject_scores_overall'].apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "### Faculty Analysis (Overall)\n\n"
    report += analysis_result['faculty_scores_overall'].apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "## Parameter-wise Feedback Analysis\n\n"

    report += "### Branch Analysis (Parameter-wise)\n\n"
    report += analysis_result['branch_scores'].reset_index().apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "### Term-Year Analysis (Parameter-wise)\n\n"
    report += analysis_result['term_year_scores'].reset_index().apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "### Semester Analysis (Parameter-wise)\n\n"
    report += analysis_result['semester_scores'].apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "### Subject Analysis (Parameter-wise)\n\n"
    report += analysis_result['subject_scores_faculty'].drop_duplicates(['Subject_Code', 'Subject_ShortForm']).apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "### Faculty Analysis (Parameter-wise)\n\n"
    report += analysis_result['faculty_scores_subject'].drop_duplicates('Faculty_Name').apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "## Misc Feedback Analysis\n\n"

    report += "### Faculty-Subject Correlation Matrix\n\n"
    
    formatted_matrix = analysis_result['correlation_matrix'].copy()
    formatted_matrix.columns = formatted_matrix.columns.map(get_faculty_initial)
    formatted_matrix.index = formatted_matrix.index.map(lambda x: f"{x[0]} ({x[1]})" if isinstance(x, tuple) else str(x))
    formatted_matrix = formatted_matrix.apply(format_float)
    
    report += formatted_matrix.to_markdown()

    return report

def get_faculty_initial(name):
    name = name.replace("Mr. ", "").replace("Ms. ", "")
    return ''.join(word[0].upper() for word in name.split())  

def generate_pdf_wkhtml(markdown_content):
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
        temp_file.write(markdown_content)

    pdf_filename = 'feedback_report_wkhtml.pdf'
    subprocess.run(['pandoc', '-s', '-o', pdf_filename, '--pdf-engine=wkhtmltopdf', '--pdf-engine-opt=--enable-local-file-access', '--css=static/css/github.css', '--toc', '-N', '--shift-heading-level-by=-1', 'temp_report_wkhtml.md'])
    os.remove('temp_report_wkhtml.md')
    return pdf_filename


def generate_pdf_latex(markdown_content):
    yaml_front_matter = '''---
title: Student Feedback Analysis Report
subtitle: EC Dept, Government Polytechnic Palanpur
margin-left: 2.5cm
margin-right: 2.5cm
margin-top: 2cm
margin-bottom: 2cm
toc: True
---
'''
    with open('temp_report.md', 'w') as temp_file:
        temp_file.write(yaml_front_matter)
        temp_file.write(markdown_content)

    pdf_filename = 'feedback_report_latex.pdf'
    subprocess.run(['pandoc', '-s', '-o', pdf_filename, '--pdf-engine=xelatex', '-N', '--shift-heading-level-by=-1', 'temp_report.md'])
    # os.remove('temp_report.md')
    return pdf_filename

def generate_report(analysis_result, original_data):
    report_markdown = generate_markdown_report(analysis_result)
    
    report_excel = 'feedback_report.xlsx'
    generate_excel_report(analysis_result, report_excel, original_data)
    
    report_pdf_wkhtml = generate_pdf_wkhtml(report_markdown)
    report_pdf_latex = generate_pdf_latex(report_markdown)
    
    zip_filename = 'feedback_report.zip'
    with zipfile.ZipFile(zip_filename, 'w') as zip_file:
        # zip_file.write(report_markdown)
        zip_file.write(report_excel)
        zip_file.write(report_pdf_wkhtml)
        zip_file.write(report_pdf_latex)

    return zip_filename

if __name__ == '__main__':
    with open('Odd_2023.csv', 'r') as file:
        original_data = file.read()
        analysis_result = analyze_feedback(original_data)
        zip_filename = generate_report(analysis_result, original_data)
        print(f"Report generated successfully. Output file: {zip_filename}")