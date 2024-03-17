import csv
from collections import defaultdict
from datetime import datetime
import matplotlib.pyplot as plt
import pandas as pd
import os
from io import StringIO 
import zipfile
import subprocess

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
    analysis['Year-Term Analysis'] = {year_term: {} for year_term in data}
    for year_term in data:
        for branch in data[year_term]:
            for semester in data[year_term][branch]:
                for subject in data[year_term][branch][semester]:
                    for faculty in data[year_term][branch][semester][subject]:
                        scores = data[year_term][branch][semester][subject][faculty]
                        analysis['Year-Term Analysis'][year_term][f"{branch} - Sem {semester}"] = calculate_average([s for sub_scores in scores for s in sub_scores])

    # Branch Analysis
    analysis['Branch Analysis'] = {}
    for year_term in data:
        for branch in data[year_term]:
            scores = [score for semester in data[year_term][branch] for subject in data[year_term][branch][semester]
                      for faculty in data[year_term][branch][semester][subject] for score in data[year_term][branch][semester][subject][faculty]]
            analysis['Branch Analysis'][branch] = {'Average scores for Q1-Q12': [calculate_average([s[i] for s in scores]) for i in range(12)],
                                                   'Overall average': calculate_average([s for sub_scores in scores for s in sub_scores])}

    # Subject Analysis
    analysis['Subject Analysis'] = {}
    for year_term in data:
        for branch in data[year_term]:
            for semester in data[year_term][branch]:
                for subject in data[year_term][branch][semester]:
                    scores = [score for faculty in data[year_term][branch][semester][subject] for score in data[year_term][branch][semester][subject][faculty]]
                    analysis['Subject Analysis'][subject] = {'Overall average': calculate_average([s for sub_scores in scores for s in sub_scores])}

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
        plt.savefig(os.path.join(output_dir, f'{category.lower()}_analysis.png'))
        plt.close()

        for param in [f'Q{i}' for i in range(1, 13)]:
            param_data = [analysis_result['Parameter-wise Analysis'][category][key][param] for key in keys]
            plt.figure(figsize=(12, 6))
            plt.bar(keys, param_data)
            plt.xlabel(category)
            plt.ylabel(f'Average Score for {param}')
            plt.title(f'Parameter-wise {category} Analysis - {param}')
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            plt.savefig(os.path.join(output_dir, f'parameter_{category.lower()}_{param}.png'))
            plt.close()

         
def generate_markdown_report(analysis_result, output_file, chart_dir):
    with open(output_file, 'w') as file:
        file.write('---\n')
        file.write("title: Student Feedback Analysis Report\n")
        file.write("subtitle: EC Dept, Government Polytechnic Palanpur\n")
        file.write("geometry: 'left=2.5cm,right=2.5cm,top=2cm,bottom=2cm'\n")
        file.write("toc: True\n")
        # file.write("mainfont: Noto Serif\n")
        # file.write("mainfont: Noto Sans Gujarati\n")
        # file.write("fontfamily: Noto Serif Gujarati\n")    
        # file.write('header-includes:\n')
        # file.write('  - |\n')
        # file.write('    ```{=latex}\n')
        # file.write('    \\usepackage{fontspec}\n')
        # file.write('    \\setmainfont{lmroman10-regular}\n')
        # file.write('    \\setmainfont{Noto Sans Gujarati}\n')
        # file.write('    ```\n')        
        # file.write("header-includes:\n")
        # file.write("  - \\usepackage{fontspec}\n")
        # file.write("  - \\usepackage{polyglossia}\n")
        # file.write("  - \\setmainlanguage{english}\n")
        # file.write("  - \\setotherlanguage{sanskrit}\n")
        # file.write("  - \\newfontfamily\\englishfont[Ligatures=TeX]{Noto Sans}\n")
        # file.write("  - \\newfontfamily\\sanskritfont[Script=Gujarati]{Noto Sans Gujarati}\n")
        file.write("---\n")
        file.write('\n')     
        
        file.write("# Student Feedback Analysis Report\n\n")
        
        file.write("## Assessment Parameters & Rating Scale\n\n")
        file.write("### Assessment Parameters\n\n")
        file.write("- **Q1 Syllabus Coverage:** Has the Teacher covered entire Syllabus as prescribed by University/ College/ Board?\n")
        file.write("- **Q2 Topics Beyond Syllabus:** Has the Teacher covered relevant topics beyond syllabus\n")
        file.write("- **Q3 Pace of Teaching:** Pace on which contents were covered\n")
        file.write("- **Q4 Practical Demo:** Support for the development of Students' skill (Practical demonstration)\n")
        file.write("- **Q5 Hands-on Training:** Support for the development of Students' skill (Hands-on training)\n")
        file.write("- **Q6 Technical Skills of Teacher:** Effectiveness of Teacher in terms of: Technical content/course content\n")
        file.write("- **Q7 Communication Skills of Teacher:** Effectiveness of Teacher in terms of: Communication skills\n")
        file.write("- **Q8 Doubt Clarification:** Clarity of expectations of students\n")
        file.write("- **Q9 Use of Teaching Tools:** Effectiveness of Teacher in terms of: Use of teaching aids\n")
        file.write("- **Q10 Motivation:** Motivation and inspiration for students to learn\n")
        file.write("- **Q11 Helpfulness of Teacher:** Willingness to offer help and advice to students\n")
        file.write("- **Q12 Student Progress Feedback:** Feedback provided on Students' progress\n\n")

        file.write("### Rating Scale\n\n")        
        file.write("Rating | Description\n")
        file.write("-------|------------\n")
        file.write("1      | Very Poor\n")
        file.write("2      | Poor\n")
        file.write("3      | Average\n")
        file.write("4      | Good\n")
        file.write("5      | Very Good\n\n")        

        file.write("## Feedback Analysis\n\n")
        file.write("### Year-Term Analysis\n\n")
        for year_term, data in analysis_result['Year-Term Analysis'].items():
            file.write(f"#### {year_term}\n\n")
            file.write("| Branch - Semester | Average Score |\n")
            file.write("|-------------------|---------------|\n")
            for key, value in data.items():
                file.write(f"| {key} | {value:.2f} |\n")
            file.write("\n")

        file.write("### Branch Analysis\n\n")
        file.write("| Branch | Overall Average | Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |\n")
        file.write("|--------|-----------------|----|----|----|----|----|----|----|----|----|----|-----|-----|\n")
        for branch, data in analysis_result['Branch Analysis'].items():
            file.write(f"| {branch} | {data['Overall average']:.2f} |")
            for score in data['Average scores for Q1-Q12']:
                file.write(f" {score:.2f} |")
            file.write("\n")
        file.write("\n")
        # file.write("![Branch Analysis](branch_analysis.png)\n\n")
        file.write(f"![Branch Analysis]({os.path.join(chart_dir, 'branch_analysis.png')})\n\n")        

        file.write("### Subject Analysis\n\n")
        file.write("| Subject | Overall Average |\n")
        file.write("|---------|------------------|\n")
        for subject, data in analysis_result['Subject Analysis'].items():
            file.write(f"| {subject} | {data['Overall average']:.2f} |\n")
        file.write("\n")
        # file.write("![Subject Analysis](subject_analysis.png)\n\n")
        file.write(f"![Subject Analysis]({os.path.join(chart_dir, 'subject_analysis.png')})\n\n")        

        file.write("### Faculty Analysis\n\n")
        for faculty, data in analysis_result['Faculty Analysis'].items():
            file.write(f"#### {faculty}\n\n")
            file.write(f"- Overall Average: {data['Overall average']:.2f}\n\n")
            file.write("| Subject | Average Score |\n")
            file.write("|---------|---------------|\n")
            for subject, average in data['Subjects'].items():
                file.write(f"| {subject} | {average:.2f} |\n")
            file.write("\n")
        # file.write("![Faculty Analysis](faculty_analysis.png)\n\n")
        file.write(f"![Faculty Analysis]({os.path.join(chart_dir, 'faculty_analysis.png')})\n\n")        

        file.write("### Parameter-wise Analysis\n\n")
        for category in ['Branch', 'Subject', 'Faculty']:
            file.write(f"#### Parameter-wise {category} Analysis\n\n")
            file.write(f"| {category} | Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |\n")
            file.write("|----------|----|----|----|----|----|----|----|----|----|-----|-----|-----|\n")
            for key, data in analysis_result['Parameter-wise Analysis'][category].items():
                file.write(f"| {key} |")
                for param in [f'Q{i}' for i in range(1, 13)]:
                    file.write(f" {data[param]:.2f} |")
                file.write("\n")
            file.write("\n")
            for param in [f'Q{i}' for i in range(1, 13)]:
                # file.write(f"![Parameter-wise {category} Analysis - {param}](parameter_{category.lower()}_{param}.png)\n\n")
                file.write(f"![Parameter-wise {category} Analysis - {param}]({os.path.join(chart_dir, f'parameter_{category.lower()}_{param}.png')})\n\n")                

def export_to_excel(analysis_result, output_file, original_data):
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    
    # Write the original data to a separate sheet
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
                df.to_excel(writer, sheet_name=f"{category} Parameter-wise")

        if sheet_name != 'Parameter-wise Analysis':
            df.to_excel(writer, sheet_name=sheet_name)

    writer._save()

def generate_pdf(markdown_file):
    pdf_filename = 'feedback_report.pdf'
    subprocess.run(['pandoc', markdown_file, '-o', pdf_filename, '--pdf-engine=wkhtmltopdf', '--pdf-engine-opt=--enable-local-file-access', '--css=github.css', '--shift-heading-level-by=-1'])
    return pdf_filename

def generate_pdf_latex(markdown_file):
    pdf_filename = 'feedback_report_latex.pdf'
    subprocess.run(['pandoc', markdown_file, '-o', pdf_filename, '--pdf-engine=xelatex', '-N', '--shift-heading-level-by=-1'])
    return pdf_filename

def generate_zip(markdown_file, charts_dir):
    zip_filename = 'feedback_report.zip'
    with zipfile.ZipFile(zip_filename, 'w') as zip_file:
        zip_file.write(markdown_file)
        for root, dirs, files in os.walk(charts_dir):
            for file in files:
                zip_file.write(os.path.join(root, file))
    return zip_filename

    
# # # Usage
# file_path = 'Odd_2023.csv'
# markdown_file = 'analysis_report.md'
# excel_file = 'analysis_report.xlsx'
# chart_dir = 'assets/imgs/charts'

# with open(file_path, 'r') as file:
#     file_content = file.read()

# analysis_result = analyze_feedback(file_content)
# generate_charts(analysis_result, chart_dir)
# generate_markdown_report(analysis_result, markdown_file, chart_dir)
# export_to_excel(analysis_result, excel_file, file_content)