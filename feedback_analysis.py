import pandas as pd
from io import StringIO 
import zipfile
import subprocess
import os

def analyze_feedback(file_content):
    data = pd.read_csv(StringIO(file_content))
    data['Subject_Code'] = data['Subject_Code'].astype(str)  # Convert Subject_Code to string 

    exclude_words = ['of', 'and', 'in', 'to', 'the', 'for', '&', 'a', 'an']  # Add more words if needed
    data['Subject_ShortForm'] = data['Subject_FullName'].apply(lambda x: ''.join(word[0].upper() for word in x.split() if word.lower() not in exclude_words))    

    subject_scores = data.groupby(['Subject_Code', 'Subject_ShortForm', 'Faculty_Name']).agg({
        **{f'Q{i}': 'mean' for i in range(1, 13)}, 'Subject_FullName': 'first'
    }).reset_index()
    subject_scores['Score'] = subject_scores[list(f'Q{i}' for i in range(1, 13))].mean(axis=1)

    faculty_scores = subject_scores.groupby('Faculty_Name').agg({
        **{f'Q{i}': 'mean' for i in range(1, 13)}, 'Score': 'mean'
    }).reset_index()
    faculty_scores['Faculty_Initial'] = faculty_scores['Faculty_Name'].apply(get_faculty_initial)

    semester_scores = data.groupby(['Year', 'Term', 'Branch', 'Sem']).agg({
        **{f'Q{i}': 'mean' for i in range(1, 13)}
    }).reset_index()

    branch_scores = semester_scores.groupby('Branch').agg({
        **{f'Q{i}': 'mean' for i in range(1, 13)}
    })

    term_year_scores = semester_scores.groupby(['Year', 'Term']).agg({
        **{f'Q{i}': 'mean' for i in range(1, 13)}
    })

    correlation_matrix = subject_scores.pivot_table(index=['Subject_Code', 'Subject_ShortForm'], columns='Faculty_Name', values='Score', aggfunc='mean')
    faculty_overall = faculty_scores.set_index('Faculty_Name')['Score']
    correlation_matrix.loc['Faculty Overall'] = faculty_overall
    correlation_matrix['Subject Overall'] = subject_scores.groupby(['Subject_Code', 'Subject_ShortForm'])['Score'].mean()
    correlation_matrix = correlation_matrix.fillna('-')

    analysis_result = {
        'subject_scores': subject_scores,
        'faculty_scores': faculty_scores,
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

    report = "# Student Feedback Analysis Report\n\n"
    
    report += "## Assessment Parameters & Rating Scale\n\n"

    report += "### Assessment Parameters\n\n"
    report += "- **Q1 Syllabus Coverage**: Has the Teacher covered the entire syllabus as prescribed by University/College/Board?\n"
    report += "- **Q2 Topics Beyond Syllabus**: Has the Teacher covered relevant topics beyond the syllabus?\n"
    report += "- **Q3 Pace of Teaching**: Pace at which contents were covered?\n"
    report += "- **Q4 Practical Demo**: Support for the development of student's skill (Practical demonstration)\n"
    report += "- **Q5 Hands-on Training**: Support for the development of student's skill (Hands-on training)\n"
    report += "- **Q6 Technical Skills of Teacher**: Effectiveness of Teacher in terms of: Technical skills\n"
    report += "- **Q7 Communication Skills of Teacher**: Effectiveness of Teacher in terms of: Communication skills\n"
    report += "- **Q8 Doubt Clarification**: Clarity of expectations of students\n"
    report += "- **Q9 Use of Teaching Tools**: Effectiveness of Teacher in terms of: Use of teaching aids\n"
    report += "- **Q10 Motivation**: Motivation and inspiration for students to learn\n"
    report += "- **Q11 Helpfulness of Teacher**: Willingness to offer help and advice to students\n"
    report += "- **Q12 Student Progress Feedback**: Feedback provided on student's progress\n\n"

    report += "### Rating Scale\n\n"
    report += "| Rating | Description |\n"
    report += "|--------|-------------|"
    report += "\n| 1      | Very Poor   |"
    report += "\n| 2      | Poor        |"
    report += "\n| 3      | Average     |"
    report += "\n| 4      | Good        |"
    report += "\n| 5      | Very Good   |\n\n"

    report += "## Feedback Analysis\n\n"

    report += "### Branch Analysis (overall)\n\n"
    report += analysis_result['branch_scores'].mean(axis=1).reset_index().rename(columns={0: 'Score'}).apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "### Term-Year Analysis (overall)\n\n"
    report += "**Term duration:**\n"
    report += "- Semester 2: 24/01/25,10/05/25\n"
    report += "- Semester 4 & 6: 18/12/24,28/04/25\n\n"
    report += analysis_result['term_year_scores'].mean(axis=1).reset_index().rename(columns={0: 'Score'}).apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "### Semester Analysis (overall)\n\n"
    report += analysis_result['semester_scores'].groupby(['Branch', 'Sem'])[list(f'Q{i}' for i in range(1, 13))].mean().mean(axis=1).reset_index().rename(columns={0: 'Score'}).apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "### Subject Analysis (overall)\n\n"
    # Get mean scores and first occurrence of Subject_FullName for each subject
    subject_scores_overall = analysis_result['subject_scores'].groupby(['Subject_Code', 'Subject_ShortForm']).agg({
        'Subject_FullName': 'first',  # Changed order to match desired output
        'Score': 'mean'
    }).reset_index()
    # Reorder columns to match desired output
    subject_scores_overall = subject_scores_overall[['Subject_Code', 'Subject_ShortForm', 'Subject_FullName', 'Score']]
    subject_scores_overall.columns = ['Subject Code', 'Subject Short Form', 'Subject Full Name', 'Score']
    report += subject_scores_overall.apply(format_float).to_markdown(index=False)
    report += "\n\n"

    report += "### Faculty Analysis (Overall)\n\n"
    faculty_scores_overall = analysis_result['faculty_scores'][['Faculty_Name', 'Faculty_Initial', 'Score']]
    report += faculty_scores_overall.apply(format_float).to_markdown(index=False)
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
    subject_scores_param = analysis_result['subject_scores'].copy()  # Create a copy to avoid modifying original
    # Add faculty initials
    subject_scores_param['Faculty_Initial'] = subject_scores_param['Faculty_Name'].apply(get_faculty_initial)
    # Format float values
    subject_scores_param = subject_scores_param.apply(format_float)
    # Select and rename columns, using Faculty_Initial instead of Faculty_Name
    subject_scores_param = subject_scores_param[['Subject_Code', 'Subject_ShortForm', 'Faculty_Initial'] + [f'Q{i}' for i in range(1, 13)] + ['Score']]
    subject_scores_param.columns = ['Subject Code', 'Subject Short Form', 'Faculty Initial'] + [f'Q{i}' for i in range(1, 13)] + ['Score']
    report += subject_scores_param.to_markdown(index=False)
    report += "\n\n"

    report += "### Faculty Analysis (Parameter-wise)\n\n"
    faculty_scores_param = analysis_result['faculty_scores'].drop(columns=['Faculty_Name']).apply(format_float)
    faculty_scores_param = faculty_scores_param[['Faculty_Initial'] + [col for col in faculty_scores_param.columns if col != 'Faculty_Initial']]
    report += faculty_scores_param.to_markdown(index=False)
    report += "\n\n"

    report += "## Misc Feedback Analysis\n\n"

    report += "### Faculty-Subject Correlation Matrix\n\n"

    formatted_matrix = analysis_result['correlation_matrix'].copy()
    formatted_matrix.columns = formatted_matrix.columns.map(get_faculty_initial)
    formatted_matrix.index = formatted_matrix.index.map(lambda x: x[0] if isinstance(x, tuple) else str(x))
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
    os.remove('temp_report.md')
    return pdf_filename

def generate_report(analysis_result, original_data):
    report_markdown = generate_markdown_report(analysis_result)
    
    report_excel = 'feedback_report.xlsx'
    generate_excel_report(analysis_result, report_excel, original_data)
    
    report_pdf_wkhtml = generate_pdf_wkhtml(report_markdown)
    report_pdf_latex = generate_pdf_latex(report_markdown)
    
    # Save the markdown content to a file
    markdown_filename = 'feedback_report.md'
    with open(markdown_filename, 'w') as markdown_file:
        markdown_file.write(report_markdown)
    
    zip_filename = 'feedback_report.zip'
    with zipfile.ZipFile(zip_filename, 'w') as zip_file:
        zip_file.write(markdown_filename)  # Write the markdown file to the ZIP archive
        zip_file.write(report_excel)
        zip_file.write(report_pdf_wkhtml)
        zip_file.write(report_pdf_latex)
    
    # # Remove the temporary markdown file
    # os.remove(markdown_filename)
    
    return zip_filename


if __name__ == '__main__':
    with open('Odd_2023.csv', 'r') as file:
        original_data = file.read()
        analysis_result = analyze_feedback(original_data)
        zip_filename = generate_report(analysis_result, original_data)
        print(f"Report generated successfully. Output file: {zip_filename}")