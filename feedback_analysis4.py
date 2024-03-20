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
    # Read the CSV data into a pandas DataFrame
    data = pd.read_csv(StringIO(file_content))

    # Calculate subject scores (faculty-wise)
    subject_scores_faculty = data.groupby(['Subject_Code', 'Subject_ShortForm', 'Faculty_Name']).agg({
        'Q1': 'mean',
        'Q2': 'mean',
        'Q3': 'mean',
        'Q4': 'mean',
        'Q5': 'mean',
        'Q6': 'mean',
        'Q7': 'mean',
        'Q8': 'mean',
        'Q9': 'mean',
        'Q10': 'mean',
        'Q11': 'mean',
        'Q12': 'mean'
    }).reset_index()
    subject_scores_faculty['Average_Score'] = subject_scores_faculty[['Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12']].mean(axis=1)

    # Calculate subject scores (overall)
    subject_scores_overall = subject_scores_faculty.groupby(['Subject_Code', 'Subject_ShortForm'])['Average_Score'].mean().reset_index()
    subject_scores_overall.columns = ['Subject_Code', 'Subject_ShortForm', 'Overall_Average']

    # Calculate faculty scores (subject-wise)
    faculty_scores_subject = subject_scores_faculty.groupby(['Faculty_Name', 'Subject_Code', 'Subject_ShortForm']).agg({
        'Q1': 'mean',
        'Q2': 'mean',
        'Q3': 'mean',
        'Q4': 'mean',
        'Q5': 'mean',
        'Q6': 'mean',
        'Q7': 'mean',
        'Q8': 'mean',
        'Q9': 'mean',
        'Q10': 'mean',
        'Q11': 'mean',
        'Q12': 'mean',
        'Average_Score': 'mean'
    }).reset_index()
    # Calculate faculty scores (overall)
    
    faculty_scores_overall = faculty_scores_subject.groupby('Faculty_Name')['Average_Score'].mean().reset_index()
    faculty_scores_overall.columns = ['Faculty_Name', 'Overall_Average']

    # Calculate semester scores
    semester_scores = data.groupby(['Year', 'Odd_Even', 'Branch', 'Sem'])[['Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12']].mean().reset_index()
    semester_scores['Branch_Semester'] = semester_scores['Branch'] + ' - ' + semester_scores['Sem'].astype(str)

    # Calculate branch scores
    branch_scores = semester_scores.groupby('Branch')[['Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12']].mean()

    # Calculate term-year scores
    term_year_scores = semester_scores.groupby(['Year', 'Odd_Even'])[['Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12']].mean()
    
    # # Create a DataFrame for the correlation matrix
    # correlation_matrix = subject_scores_faculty.pivot_table(index=['Subject_Code', 'Subject_ShortForm'], columns='Faculty_Name', values='Average_Score', aggfunc='mean')
    # correlation_matrix['Overall'] = subject_scores_overall['Overall_Average']
    # correlation_matrix.loc['Overall'] = faculty_scores_overall.set_index('Faculty_Name')['Overall_Average']

    # Create a DataFrame for the correlation matrix
    correlation_matrix = subject_scores_faculty.pivot_table(index=['Subject_Code', 'Subject_ShortForm'], columns='Faculty_Name', values='Average_Score', aggfunc='mean')

    # Create a dictionary to map faculty names to their initials
    faculty_initials = {
        'Mr. L K Patel': 'LKP',
        'Mr. M J Dabgar': 'MJD',
        'Mr. M J Vadhwania': 'MJV',
        'Mr. N J Chauhan': 'NJC',
        'Mr. R C Parmar': 'RCP',
        'Mr. R N Patel': 'RNP',
        'Mr. S J Chauhan': 'SJC',
        'Mr. S P Joshiara': 'SPJ',
        'Ms. M K Pedhadiya': 'MKP'
    }

    # Rename the columns in the correlation matrix with faculty initials
    correlation_matrix.columns = [faculty_initials.get(col, col) for col in correlation_matrix.columns]

    # Add 'Overall' column to the correlation matrix
    correlation_matrix['Overall'] = subject_scores_overall.set_index(['Subject_Code', 'Subject_ShortForm'])['Overall_Average']

    # Create a new index for the overall_row with the same index as correlation_matrix
    new_index = pd.MultiIndex.from_product([['Overall'], correlation_matrix.columns], names=['Subject', 'Faculty'])

    # Add 'Overall' row to the correlation matrix
    overall_row = pd.Series(faculty_scores_overall.set_index('Faculty_Name')['Overall_Average'], index=new_index)
    correlation_matrix = pd.concat([correlation_matrix, overall_row.to_frame().T])

    # Fill 'nan' values with '-'
    correlation_matrix.fillna('-', inplace=True)
     
    # Prepare the analysis results
    analysis_result = {
        'subject_scores_faculty': subject_scores_faculty,
        'subject_scores_overall': subject_scores_overall,
        'faculty_scores_subject': faculty_scores_subject,
        'faculty_scores_overall': faculty_scores_overall,
        'semester_scores': semester_scores,
        'branch_scores': branch_scores,
        'term_year_scores': term_year_scores,
        'correlation_matrix': correlation_matrix, 
        'faculty_initials': faculty_initials
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

def generate_markdown_report(analysis_result, markdown_file):
    subject_scores_faculty = analysis_result['subject_scores_faculty']
    subject_scores_overall = analysis_result['subject_scores_overall']
    faculty_scores_subject = analysis_result['faculty_scores_subject']
    faculty_scores_overall = analysis_result['faculty_scores_overall']
    semester_scores = analysis_result['semester_scores']
    branch_scores = analysis_result['branch_scores']
    term_year_scores = analysis_result['term_year_scores']
    correlation_matrix = analysis_result['correlation_matrix']
    faculty_initials = analysis_result['faculty_initials']

    report = "## Feedback Analysis\n\n"

    report += "## Overall Feedback Analysis\n\n"
    report += "| Branch Score | Term-Year Score | Semester Score | Subject Score |\n"
    report += "| --- | --- | --- | --- |\n"
    report += f"| {branch_scores.mean().mean():.2f} | {term_year_scores.mean().mean():.2f} | {semester_scores[['Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12']].mean().mean():.2f} | {subject_scores_overall['Overall_Average'].mean():.2f} |\n\n"

    report += "### Branch Analysis (overall)\n\n"
    report += "| Branch | Average Score |\n"
    report += "|--------|---------------|\n"
    for branch, row in branch_scores.iterrows():
        report += f"| {branch} | {row.mean():.2f} |\n"
    report += "\n"

    report += "### Term-Year Analysis (overall)\n\n"
    report += "| Term-Year | Overall |\n"
    report += "| --- | --- |\n"
    for term_year, row in term_year_scores.iterrows():
        report += f"| {term_year[0]}-{term_year[1]} | {row.mean():.2f} |\n"
    report += "\n"

    report += "### Semester Analysis (overall)\n\n"
    report += "| Branch - Semester | Average Score |\n"
    report += "|----------|---------------|\n"
    for _, row in semester_scores.iterrows():
        report += f"| {row['Branch_Semester']} | {row[['Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12']].mean():.2f} |\n"
    report += "\n"

    report += "### Subject Analysis (overall)\n\n"
    report += "| Subject | Overall Average |\n"
    report += "|---------|------------------|\n"
    for _, row in subject_scores_overall.iterrows():
        report += f"| {row['Subject_Code']} ({row['Subject_ShortForm']}) | {row['Overall_Average']:.2f} |\n"
    report += "\n"

    report += "### Faculty Analysis (Overall)\n\n"
    report += "| Faculty | Overall Average |\n"
    report += "|---------|------------------|\n"
    for _, row in faculty_scores_overall.iterrows():
        report += f"| {row['Faculty_Name']} | {row['Overall_Average']:.2f} |\n"
    report += "\n"

    report += "## Parameter-wise Feedback Analysis\n\n"

    report += "### Branch Analysis (Parameter-wise)\n\n"
    report += "| Branch | Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |\n"
    report += "|--------|---|---|---|---|---|---|---|---|---|---|---|---|\n"
    for branch, row in branch_scores.iterrows():
        report += f"| {branch} | {row['Q1']:.2f} | {row['Q2']:.2f} | {row['Q3']:.2f} | {row['Q4']:.2f} | {row['Q5']:.2f} | {row['Q6']:.2f} | {row['Q7']:.2f} | {row['Q8']:.2f} | {row['Q9']:.2f} | {row['Q10']:.2f} | {row['Q11']:.2f} | {row['Q12']:.2f} |\n"
    report += "\n"

    report += "### Term-Year Analysis (Parameter-wise)\n\n"
    report += "| Term-Year | Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |\n"
    report += "|-----------|---|---|---|---|---|---|---|---|---|---|---|---|\n"
    for term_year, row in term_year_scores.iterrows():
        report += f"| {term_year[0]}-{term_year[1]} | {row['Q1']:.2f} | {row['Q2']:.2f} | {row['Q3']:.2f} | {row['Q4']:.2f} | {row['Q5']:.2f} | {row['Q6']:.2f} | {row['Q7']:.2f} | {row['Q8']:.2f} | {row['Q9']:.2f} | {row['Q10']:.2f} | {row['Q11']:.2f} | {row['Q12']:.2f} |\n"
    report += "\n"

    report += "### Semester Analysis (Parameter-wise)\n\n"
    report += "| Branch - Semester | Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |\n"
    report += "|-------------------|---|---|---|---|---|---|---|---|---|---|---|---|\n"
    for _, row in semester_scores.iterrows():
        report += f"| {row['Branch_Semester']} | {row['Q1']:.2f} | {row['Q2']:.2f} | {row['Q3']:.2f} | {row['Q4']:.2f} | {row['Q5']:.2f} | {row['Q6']:.2f} | {row['Q7']:.2f} | {row['Q8']:.2f} | {row['Q9']:.2f} | {row['Q10']:.2f} | {row['Q11']:.2f} | {row['Q12']:.2f} |\n"
    report += "\n"

    report += "### Subject Analysis (Parameter-wise)\n\n"
    report += "| Subject | Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |\n"
    report += "|---------|---|---|---|---|---|---|---|---|---|---|---|---|\n"
    for _, row in subject_scores_faculty.drop_duplicates(['Subject_Code', 'Subject_ShortForm']).iterrows():
        report += f"| {row['Subject_Code']} ({row['Subject_ShortForm']}) | {row['Q1']:.2f} | {row['Q2']:.2f} | {row['Q3']:.2f} | {row['Q4']:.2f} | {row['Q5']:.2f} | {row['Q6']:.2f} | {row['Q7']:.2f} | {row['Q8']:.2f} | {row['Q9']:.2f} | {row['Q10']:.2f} | {row['Q11']:.2f} | {row['Q12']:.2f} |\n"
    report += "\n"

    report += "### Faculty Analysis (Parameter-wise)\n\n"
    report += "| Faculty | Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |\n"
    report += "|---------|---|---|---|---|---|---|---|---|---|---|---|---|\n"
    for _, row in faculty_scores_subject.drop_duplicates('Faculty_Name').iterrows():
        report += f"| {row['Faculty_Name']} | {row['Q1']:.2f} | {row['Q2']:.2f} | {row['Q3']:.2f} | {row['Q4']:.2f} | {row['Q5']:.2f} | {row['Q6']:.2f} | {row['Q7']:.2f} | {row['Q8']:.2f} | {row['Q9']:.2f} | {row['Q10']:.2f} | {row['Q11']:.2f} | {row['Q12']:.2f} |\n"
    report += "\n"

    report += "## Misc Feedback Analysis\n\n"

    report += "### Subject Analysis (Faculty-wise)\n\n"
    for subject_code, subject_data in subject_scores_faculty.groupby(['Subject_Code', 'Subject_ShortForm']):
        report += f"#### {subject_code[0]} ({subject_code[1]})\n\n"
        report += f"- Overall Average: {subject_data['Average_Score'].mean():.2f}\n\n"
        report += "| Faculty | Average Score |\n"
        report += "|---------|---------------|\n"
        for _, row in subject_data.iterrows():
            report += f"| {row['Faculty_Name']} | {row['Average_Score']:.2f} |\n"
        report += "\n"

    report += "### Faculty Analysis (Subject-wise)\n\n"
    for faculty_name, faculty_data in faculty_scores_subject.groupby('Faculty_Name'):
        report += f"#### {faculty_name}\n\n"
        report += f"- Overall Average: {faculty_data['Average_Score'].mean():.2f}\n\n"
        report += "| Subject | Average Score |\n"
        report += "|---------|---------------|\n"
        for _, row in faculty_data.iterrows():
            report += f"| {row['Subject_Code']} ({row['Subject_ShortForm']}) | {row['Average_Score']:.2f} |\n"
        report += "\n"

    report += "### Faculty-Subject Correlation Matrix\n\n"
    report += "| Subject | " + " | ".join(correlation_matrix.columns) + " |\n"
    report += "|---------|" + "--|" * (len(correlation_matrix.columns)) + "\n"
    for subject, row in correlation_matrix.iterrows():
        if subject == 'Overall':
            report += "| Overall |"
        else:
            report += f"| {subject[0]} ({subject[1]}) |"
        for score in row:
            report += f" {score} |"
        report += "\n"
    report += "\n"

    report += "Faculty Initials:\n\n"
    for name, initial in faculty_initials.items():
        report += f"- {initial}: {name}\n"
    report += "\n"

    # # Save the report to a markdown file
    # with open(markdown_file, 'w') as file:
    #     file.write(report)

    # report += "### Faculty-Subject Correlation Matrix\n\n"
    # report += "| Subject | " + " | ".join(correlation_matrix.columns) + " |\n"
    # report += "|---------|" + "--|" * (len(correlation_matrix.columns)) + "\n"
    # for subject, row in correlation_matrix.iterrows():
    #     if subject == 'Overall':
    #         report += "| Overall |"
    #     else:
    #         report += f"| {subject[0]} ({subject[1]}) |"
    #     for score in row:
    #         if isinstance(score, (int, float)):
    #             report += f" {score:.2f} |"
    #         else:
    #             report += f" {score} |"
    #     report += "\n"
    # report += "\n"

    # report += "Faculty Initials:\n"
    # for name, initial in faculty_initials.items():
    #     report += f"- {initial}: {name}\n"
    # report += "\n"

    # # Save the report to a markdown file
    # with open(markdown_file, 'w') as file:
    #     file.write(report)     
            
    # report += "### Faculty-Subject Correlation Matrix\n\n"
    # report += "| Subject | " + " | ".join(correlation_matrix.columns) + " |\n"
    # report += "|---------|" + "--|" * (len(correlation_matrix.columns)) + "\n"
    # for subject, row in correlation_matrix.iterrows():
    #     if subject != 'Overall':  # Skip the 'Overall' row
    #         report += f"| {subject[0]} ({subject[1]}) |"
    #         for score in row:
    #             if isinstance(score, (int, float)):
    #                 report += f" {score:.2f} |"
    #             else:
    #                 report += f" {score} |"
    #         report += "\n"

    # # Add the faculty overall scores row
    # report += "| Overall |"
    # for faculty_initial in correlation_matrix.columns:
    #     if faculty_initial == 'Overall':
    #         report += " - |"
    #     else:
    #         faculty_name = [name for name, initial in faculty_initials.items() if initial == faculty_initial][0]
    #         faculty_score = faculty_scores_overall.set_index('Faculty_Name').loc[faculty_name, 'Overall_Average']
    #         report += f" {faculty_score:.2f} |"
    # report += "\n"

    # report += "\n"
    # report += "Faculty Initials:\n"
    # for name, initial in faculty_initials.items():
    #     report += f"- {initial}: {name}\n"
    # report += "\n"

    # # Save the report to a markdown file
    # with open(markdown_file, 'w') as file:
    #     file.write(report)

    # report += "### Faculty-Subject Correlation Matrix\n\n"

    # # Convert column MultiIndex to strings
    # column_strings = [' '.join(str(col_item) for col_item in col) for col in correlation_matrix.columns.values]
    # report += "| Subject | " + " | ".join(column_strings) + " |\n"

    # report += "|---------|" + "--|" * (len(correlation_matrix.columns)) + "\n"
    # for subject, row in correlation_matrix.iterrows():
    #     if isinstance(subject, tuple):
    #         subject_str = ' '.join(str(item) for item in subject)
    #         report += f"| {subject_str} |"
    #     else:
    #         report += f"| {subject} |"
    #     for score in row:
    #         if isinstance(score, (int, float)):
    #             report += f" {score:.2f} |"
    #         else:
    #             report += f" {score} |"
    #     report += "\n"

    # report += "\n"
    # report += "Faculty Initials:\n"
    # for name, initial in faculty_initials.items():
    #     report += f"- {initial}: {name}\n"
    # report += "\n"

    # # Save the report to a markdown file
    # with open(markdown_file, 'w') as file:
    #     file.write(report)

    report += "### Faculty-Subject Correlation Matrix\n\n"
    report += "| Subject | " + " | ".join(str(col) for col in correlation_matrix.columns) + " |\n"
    report += "|---------|" + "--|" * (len(correlation_matrix.columns)) + "\n"
    for index, row in correlation_matrix.iterrows():
        if isinstance(index, tuple):
            subject_code, subject_short_form = index
            report += f"| {subject_code} ({subject_short_form}) |"
        else:
            report += "| Overall |"
        
        for score in row:
            if isinstance(score, (int, float)):
                report += f" {score:.2f} |"
            else:
                report += f" {score} |"
        report += "\n"
    report += "\n"

    report += "Faculty Initials:\n"
    for name, initial in faculty_initials.items():
        report += f"- {initial}: {name}\n"
    report += "\n"

    # ...

    # Save the report to a markdown file
    with open(markdown_file, 'w') as file:
        file.write(report)

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
    
    # excel_file = 'feedback_report.xlsx'
    # generate_excel_report(analysis_result, excel_file, original_data)
    
    pdf_wkhtml = generate_pdf_wkhtml(markdown_file)
    pdf_latex = generate_pdf_latex(markdown_file)
    
    zip_filename = 'feedback_report.zip'
    with zipfile.ZipFile(zip_filename, 'w') as zip_file:
        zip_file.write(markdown_file)
        # zip_file.write(excel_file)
        zip_file.write(pdf_wkhtml)
        zip_file.write(pdf_latex)
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                zip_file.write(os.path.join(root, file))    