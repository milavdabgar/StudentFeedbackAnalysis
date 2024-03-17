import csv
from collections import defaultdict
from datetime import datetime
import matplotlib.pyplot as plt

def calculate_average(scores):
    return sum(scores) / len(scores)

def analyze_feedback(file_path):
    data = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(list)))))

    with open(file_path, 'r') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            term = row['Odd_Even']
            branch = row['Branch']
            semester = row['Sem']
            subject = f"{row['Subject_ShortForm']} ({row['Subject_Code']})"
            faculty = row['Faculty_Name']
            response_count = int(row['Responce_Count'])

            scores = [int(row[f'Q{i}']) for i in range(1, 13)]
            data[term][branch][semester][subject][faculty].append((scores, response_count))

    analysis = {}

    # Term Analysis
    terms = list(data.keys())
    analysis['Term Analysis'] = f"The data is only for the {', '.join(terms)} term(s), so no term-wise comparison can be made."

    # Branch Analysis
    analysis['Branch Analysis'] = {}
    for term in data:
        for branch in data[term]:
            if branch not in analysis['Branch Analysis']:
                analysis['Branch Analysis'][branch] = {}
            scores = [score for semester in data[term][branch] for subject in data[term][branch][semester]
                      for faculty in data[term][branch][semester][subject] for score, _ in data[term][branch][semester][subject][faculty]]
            analysis['Branch Analysis'][branch]['Average scores for Q1-Q12'] = [calculate_average([s[i] for s in scores]) for i in range(12)]
            analysis['Branch Analysis'][branch]['Overall average'] = calculate_average([s for sub_scores in scores for s in sub_scores])

    # Semester Analysis
    analysis['Semester Analysis'] = {}
    for term in data:
        for branch in data[term]:
            for semester in data[term][branch]:
                scores = [score for subject in data[term][branch][semester] for faculty in data[term][branch][semester][subject]
                          for score, _ in data[term][branch][semester][subject][faculty]]
                analysis['Semester Analysis'][f"Sem {semester} ({branch})"] = {'Overall average': calculate_average([s for sub_scores in scores for s in sub_scores])}

    # Subject Analysis
    analysis['Subject Analysis'] = {}
    for term in data:
        for branch in data[term]:
            for semester in data[term][branch]:
                for subject in data[term][branch][semester]:
                    if subject not in analysis['Subject Analysis']:
                        analysis['Subject Analysis'][subject] = {}
                    scores = [score for faculty in data[term][branch][semester][subject] for score, _ in data[term][branch][semester][subject][faculty]]
                    analysis['Subject Analysis'][subject]['Overall average'] = calculate_average([s for sub_scores in scores for s in sub_scores])

    # Faculty Analysis
    analysis['Faculty Analysis'] = {}
    for term in data:
        for branch in data[term]:
            for semester in data[term][branch]:
                for subject in data[term][branch][semester]:
                    for faculty in data[term][branch][semester][subject]:
                        if faculty not in analysis['Faculty Analysis']:
                            analysis['Faculty Analysis'][faculty] = {}
                            analysis['Faculty Analysis'][faculty]['Subjects'] = {}
                        scores = [score for score, _ in data[term][branch][semester][subject][faculty]]
                        analysis['Faculty Analysis'][faculty]['Subjects'][subject] = calculate_average([s for sub_scores in scores for s in sub_scores])
                        if 'Overall average' not in analysis['Faculty Analysis'][faculty]:
                            analysis['Faculty Analysis'][faculty]['Overall average'] = []
                        analysis['Faculty Analysis'][faculty]['Overall average'].extend([s for sub_scores in scores for s in sub_scores])

    for faculty in analysis['Faculty Analysis']:
        analysis['Faculty Analysis'][faculty]['Overall average'] = calculate_average(analysis['Faculty Analysis'][faculty]['Overall average'])

    # Response Count Analysis
    analysis['Response Count Analysis'] = {}
    for term in data:
        for branch in data[term]:
            for semester in data[term][branch]:
                for subject in data[term][branch][semester]:
                    for faculty in data[term][branch][semester][subject]:
                        response_count = sum(count for _, count in data[term][branch][semester][subject][faculty])
                        if subject not in analysis['Response Count Analysis']:
                            analysis['Response Count Analysis'][subject] = {}
                        analysis['Response Count Analysis'][subject][faculty] = response_count

    # Top Subjects
    subject_averages = {subject: subject_data['Overall average'] for subject, subject_data in analysis['Subject Analysis'].items()}
    analysis['Top Subjects'] = sorted(subject_averages.items(), key=lambda x: x[1], reverse=True)[:5]

    # Top Faculty
    faculty_averages = {faculty: faculty_data['Overall average'] for faculty, faculty_data in analysis['Faculty Analysis'].items()}
    analysis['Top Faculty'] = sorted(faculty_averages.items(), key=lambda x: x[1], reverse=True)[:5]

    # Parameter-wise Analysis
    analysis['Parameter-wise Analysis'] = {}

    # Parameter-wise Branch Analysis
    analysis['Parameter-wise Analysis']['Branch'] = {}
    for term in data:
        for branch in data[term]:
            if branch not in analysis['Parameter-wise Analysis']['Branch']:
                analysis['Parameter-wise Analysis']['Branch'][branch] = {f'Q{i}': [] for i in range(1, 13)}
            for semester in data[term][branch]:
                for subject in data[term][branch][semester]:
                    for faculty in data[term][branch][semester][subject]:
                        for scores, _ in data[term][branch][semester][subject][faculty]:
                            for i in range(12):
                                analysis['Parameter-wise Analysis']['Branch'][branch][f'Q{i+1}'].append(scores[i])

    for branch in analysis['Parameter-wise Analysis']['Branch']:
        for param in analysis['Parameter-wise Analysis']['Branch'][branch]:
            analysis['Parameter-wise Analysis']['Branch'][branch][param] = calculate_average(analysis['Parameter-wise Analysis']['Branch'][branch][param])

    # Parameter-wise Semester Analysis
    analysis['Parameter-wise Analysis']['Semester'] = {}
    for term in data:
        for branch in data[term]:
            for semester in data[term][branch]:
                if f"Sem {semester} ({branch})" not in analysis['Parameter-wise Analysis']['Semester']:
                    analysis['Parameter-wise Analysis']['Semester'][f"Sem {semester} ({branch})"] = {f'Q{i}': [] for i in range(1, 13)}
                for subject in data[term][branch][semester]:
                    for faculty in data[term][branch][semester][subject]:
                        for scores, _ in data[term][branch][semester][subject][faculty]:
                            for i in range(12):
                                analysis['Parameter-wise Analysis']['Semester'][f"Sem {semester} ({branch})"][f'Q{i+1}'].append(scores[i])

    for semester in analysis['Parameter-wise Analysis']['Semester']:
        for param in analysis['Parameter-wise Analysis']['Semester'][semester]:
            analysis['Parameter-wise Analysis']['Semester'][semester][param] = calculate_average(analysis['Parameter-wise Analysis']['Semester'][semester][param])

    # Parameter-wise Subject Analysis
    analysis['Parameter-wise Analysis']['Subject'] = {}
    for term in data:
        for branch in data[term]:
            for semester in data[term][branch]:
                for subject in data[term][branch][semester]:
                    if subject not in analysis['Parameter-wise Analysis']['Subject']:
                        analysis['Parameter-wise Analysis']['Subject'][subject] = {f'Q{i}': [] for i in range(1, 13)}
                    for faculty in data[term][branch][semester][subject]:
                        for scores, _ in data[term][branch][semester][subject][faculty]:
                            for i in range(12):
                                analysis['Parameter-wise Analysis']['Subject'][subject][f'Q{i+1}'].append(scores[i])

    for subject in analysis['Parameter-wise Analysis']['Subject']:
        for param in analysis['Parameter-wise Analysis']['Subject'][subject]:
            analysis['Parameter-wise Analysis']['Subject'][subject][param] = calculate_average(analysis['Parameter-wise Analysis']['Subject'][subject][param])

    # Parameter-wise Faculty Analysis
    analysis['Parameter-wise Analysis']['Faculty'] = {}
    for term in data:
        for branch in data[term]:
            for semester in data[term][branch]:
                for subject in data[term][branch][semester]:
                    for faculty in data[term][branch][semester][subject]:
                        if faculty not in analysis['Parameter-wise Analysis']['Faculty']:
                            analysis['Parameter-wise Analysis']['Faculty'][faculty] = {f'Q{i}': [] for i in range(1, 13)}
                        for scores, _ in data[term][branch][semester][subject][faculty]:
                            for i in range(12):
                                analysis['Parameter-wise Analysis']['Faculty'][faculty][f'Q{i+1}'].append(scores[i])

    for faculty in analysis['Parameter-wise Analysis']['Faculty']:
        for param in analysis['Parameter-wise Analysis']['Faculty'][faculty]:
            analysis['Parameter-wise Analysis']['Faculty'][faculty][param] = calculate_average(analysis['Parameter-wise Analysis']['Faculty'][faculty][param])

    return analysis, data


def generate_charts(analysis_result):
    # Branch Analysis Chart
    branches = list(analysis_result['Branch Analysis'].keys())
    branch_averages = [data['Overall average'] for data in analysis_result['Branch Analysis'].values()]
    plt.figure(figsize=(10, 6))
    plt.bar(branches, branch_averages)
    plt.xlabel('Branch')
    plt.ylabel('Overall Average')
    plt.title('Branch Analysis')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('branch_analysis.png')
    plt.close()

    # Semester Analysis Chart
    semesters = list(analysis_result['Semester Analysis'].keys())
    semester_averages = [data['Overall average'] for data in analysis_result['Semester Analysis'].values()]
    plt.figure(figsize=(10, 6))
    plt.bar(semesters, semester_averages)
    plt.xlabel('Semester')
    plt.ylabel('Overall Average')
    plt.title('Semester Analysis')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('semester_analysis.png')
    plt.close()

    # Subject Analysis Chart
    subjects = list(analysis_result['Subject Analysis'].keys())
    subject_averages = [data['Overall average'] for data in analysis_result['Subject Analysis'].values()]
    plt.figure(figsize=(12, 6))
    plt.bar(subjects, subject_averages)
    plt.xlabel('Subject')
    plt.ylabel('Overall Average')
    plt.title('Subject Analysis')
    plt.xticks(rotation=90)
    plt.tight_layout()
    plt.savefig('subject_analysis.png')
    plt.close()

    # Faculty Analysis Chart
    faculties = list(analysis_result['Faculty Analysis'].keys())
    faculty_averages = [data['Overall average'] for data in analysis_result['Faculty Analysis'].values()]
    plt.figure(figsize=(12, 6))
    plt.bar(faculties, faculty_averages)
    plt.xlabel('Faculty')
    plt.ylabel('Overall Average')
    plt.title('Faculty Analysis')
    plt.xticks(rotation=90)
    plt.tight_layout()
    plt.savefig('faculty_analysis.png')
    plt.close()

    # Parameter-wise Branch Analysis Chart
    branches = list(analysis_result['Parameter-wise Analysis']['Branch'].keys())
    params = list(analysis_result['Parameter-wise Analysis']['Branch'][branches[0]].keys())
    param_data = {param: [analysis_result['Parameter-wise Analysis']['Branch'][branch][param] for branch in branches] for param in params}
    fig, ax = plt.subplots(figsize=(10, 6))
    x = range(len(branches))
    width = 0.8 / len(params)
    for i, param in enumerate(params):
        ax.bar([xpos + i * width for xpos in x], param_data[param], width, label=param)
    ax.set_xlabel('Branch')
    ax.set_ylabel('Average Score')
    ax.set_title('Parameter-wise Branch Analysis')
    ax.set_xticks([xpos + (len(params) - 1) * width / 2 for xpos in x])
    ax.set_xticklabels(branches, rotation=45)
    ax.legend()
    plt.tight_layout()
    plt.savefig('parameter_branch_analysis.png')
    plt.close()

    # Parameter-wise Semester Analysis Chart
    semesters = list(analysis_result['Parameter-wise Analysis']['Semester'].keys())
    params = list(analysis_result['Parameter-wise Analysis']['Semester'][semesters[0]].keys())
    param_data = {param: [analysis_result['Parameter-wise Analysis']['Semester'][semester][param] for semester in semesters] for param in params}
    fig, ax = plt.subplots(figsize=(12, 6))
    x = range(len(semesters))
    width = 0.8 / len(params)
    for i, param in enumerate(params):
        ax.bar([xpos + i * width for xpos in x], param_data[param], width, label=param)
    ax.set_xlabel('Semester')
    ax.set_ylabel('Average Score')
    ax.set_title('Parameter-wise Semester Analysis')
    ax.set_xticks([xpos + (len(params) - 1) * width / 2 for xpos in x])
    ax.set_xticklabels(semesters, rotation=90)
    ax.legend()
    plt.tight_layout()
    plt.savefig('parameter_semester_analysis.png')
    plt.close()

    # Parameter-wise Subject Analysis Chart
    subjects = list(analysis_result['Parameter-wise Analysis']['Subject'].keys())
    params = list(analysis_result['Parameter-wise Analysis']['Subject'][subjects[0]].keys())
    param_data = {param: [analysis_result['Parameter-wise Analysis']['Subject'][subject][param] for subject in subjects] for param in params}
    fig, ax = plt.subplots(figsize=(12, 8))
    x = range(len(subjects))
    width = 0.8 / len(params)
    for i, param in enumerate(params):
        ax.bar([xpos + i * width for xpos in x], param_data[param], width, label=param)
    ax.set_xlabel('Subject')
    ax.set_ylabel('Average Score')
    ax.set_title('Parameter-wise Subject Analysis')
    ax.set_xticks([xpos + (len(params) - 1) * width / 2 for xpos in x])
    ax.set_xticklabels(subjects, rotation=90)
    ax.legend()
    plt.tight_layout()
    plt.savefig('parameter_subject_analysis.png')
    plt.close()

    # Parameter-wise Faculty Analysis Chart
    faculties = list(analysis_result['Parameter-wise Analysis']['Faculty'].keys())
    params = list(analysis_result['Parameter-wise Analysis']['Faculty'][faculties[0]].keys())
    param_data = {param: [analysis_result['Parameter-wise Analysis']['Faculty'][faculty][param] for faculty in faculties] for param in params}
    fig, ax = plt.subplots(figsize=(12, 8))
    x = range(len(faculties))
    width = 0.8 / len(params)
    for i, param in enumerate(params):
        ax.bar([xpos + i * width for xpos in x], param_data[param], width, label=param)
    ax.set_xlabel('Faculty')
    ax.set_ylabel('Average Score')
    ax.set_title('Parameter-wise Faculty Analysis')
    ax.set_xticks([xpos + (len(params) - 1) * width / 2 for xpos in x])
    ax.set_xticklabels(faculties, rotation=90)
    ax.legend()
    plt.tight_layout()
    plt.savefig('parameter_faculty_analysis.png')
    plt.close()

def generate_markdown_report(analysis_result, data, output_file):
    with open(output_file, 'w') as file:
        file.write("# Student Feedback Analysis Report\n\n")
        file.write("## EC Dept, Government Polytechnic Palanpur\n\n")

        current_date = datetime.now().strftime("%B %d, %Y")
        file.write(f"*Generated on: {current_date}*\n\n")

        file.write("### Branch Rating Details\n\n")
        for term in data:
            for branch in data[term]:
                file.write(f"- {branch}: {analysis_result['Branch Analysis'][branch]['Overall average']:.2f}\n")
                file.write(f"  - {term} Term: {analysis_result['Branch Analysis'][branch]['Overall average']:.2f}\n")
                for semester in data[term][branch]:
                    file.write(f"    - Sem{semester} ({branch}): {analysis_result['Semester Analysis'][f'Sem {semester} ({branch})']['Overall average']:.2f}\n")
                    for subject in data[term][branch][semester]:
                        file.write(f"      - {subject}: {analysis_result['Subject Analysis'][subject]['Overall average']:.2f}\n")
                        for faculty in data[term][branch][semester][subject]:
                            scores = [score for score, _ in data[term][branch][semester][subject][faculty]]
                            file.write(f"        - {faculty}: {calculate_average([s for sub_scores in scores for s in sub_scores]):.2f}\n")

        file.write("\n### Faculty Rating Details\n\n")
        for faculty, faculty_data in analysis_result['Faculty Analysis'].items():
            file.write(f"- {faculty}: {faculty_data['Overall average']:.2f}\n")
            for subject, rating in faculty_data['Subjects'].items():
                file.write(f"  - {subject}: {rating:.2f}\n")

        file.write("\n### Analysis Tables\n\n")

        file.write("#### Branch Analysis\n\n")
        file.write("| Branch | Overall Average |\n")
        file.write("|--------|----------------|\n")
        for branch, branch_data in analysis_result['Branch Analysis'].items():
            file.write(f"| {branch} | {branch_data['Overall average']:.2f} |\n")
        file.write("\n![Branch Analysis](branch_analysis.png)\n\n")

        file.write("\n#### Semester Analysis\n\n")
        file.write("| Semester | Branch | Overall Average |\n")
        file.write("|----------|--------|----------------|\n")
        for semester, semester_data in analysis_result['Semester Analysis'].items():
            branch = semester.split('(')[1].split(')')[0]
            file.write(f"| {semester} | {branch} | {semester_data['Overall average']:.2f} |\n")
        file.write("\n![Semester Analysis](semester_analysis.png)\n\n")

        file.write("\n#### Subject Analysis\n\n")
        file.write("| Subject | Overall Average |\n")
        file.write("|---------|----------------|\n")
        for subject, subject_data in analysis_result['Subject Analysis'].items():
            file.write(f"| {subject} | {subject_data['Overall average']:.2f} |\n")
        file.write("\n![Subject Analysis](subject_analysis.png)\n\n")

        file.write("\n#### Faculty Analysis\n\n")
        file.write("| Faculty | Overall Average |\n")
        file.write("|---------|----------------|\n")
        for faculty, faculty_data in analysis_result['Faculty Analysis'].items():
            file.write(f"| {faculty} | {faculty_data['Overall average']:.2f} |\n")
        file.write("\n![Faculty Analysis](faculty_analysis.png)\n\n")

        file.write("\n#### Faculty-Subject Analysis\n\n")
        file.write("| Faculty | Subject | Average |\n")
        file.write("|---------|---------|--------|\n")
        for faculty, faculty_data in analysis_result['Faculty Analysis'].items():
            for subject, rating in faculty_data['Subjects'].items():
                file.write(f"| {faculty} | {subject} | {rating:.2f} |\n")

        file.write("\n#### Response Count Analysis\n\n")
        file.write("| Subject | Faculty | Response Count |\n")
        file.write("|---------|---------|---------------|\n")
        for subject, faculty_counts in analysis_result['Response Count Analysis'].items():
            for faculty, count in faculty_counts.items():
                file.write(f"| {subject} | {faculty} | {count} |\n")

        file.write("\n#### Top Subjects\n\n")
        file.write("| Subject | Overall Average |\n")
        file.write("|---------|----------------|\n")
        for subject, average in analysis_result['Top Subjects']:
            file.write(f"| {subject} | {average:.2f} |\n")

        file.write("\n#### Top Faculty\n\n")
        file.write("| Faculty | Overall Average |\n")
        file.write("|---------|----------------|\n")
        for faculty, average in analysis_result['Top Faculty']:
            file.write(f"| {faculty} | {average:.2f} |\n")

        file.write("\n### Parameter-wise Analysis\n\n")

        file.write("#### Parameter-wise Branch Analysis\n\n")
        file.write("| Branch | " + " | ".join(f"Q{i}" for i in range(1, 13)) + " |\n")
        file.write("|" + "---|" * 13 + "\n")
        for branch, param_data in analysis_result['Parameter-wise Analysis']['Branch'].items():
            row = f"| {branch} |"
            for param, average in param_data.items():
                row += f" {average:.2f} |"
            file.write(row + "\n")
        file.write("\n![Parameter-wise Branch Analysis](parameter_branch_analysis.png)\n\n")

        file.write("\n#### Parameter-wise Semester Analysis\n\n")
        file.write("| Semester | " + " | ".join(f"Q{i}" for i in range(1, 13)) + " |\n")
        file.write("|" + "---|" * 13 + "\n")
        for semester, param_data in analysis_result['Parameter-wise Analysis']['Semester'].items():
            row = f"| {semester} |"
            for param, average in param_data.items():
                row += f" {average:.2f} |"
            file.write(row + "\n")
        file.write("\n![Parameter-wise Semester Analysis](parameter_semester_analysis.png)\n\n")

        file.write("\n#### Parameter-wise Subject Analysis\n\n")
        file.write("| Subject | " + " | ".join(f"Q{i}" for i in range(1, 13)) + " |\n")
        file.write("|" + "---|" * 13 + "\n")
        for subject, param_data in analysis_result['Parameter-wise Analysis']['Subject'].items():
            row = f"| {subject} |"
            for param, average in param_data.items():
                row += f" {average:.2f} |"
            file.write(row + "\n")
        file.write("\n![Parameter-wise Subject Analysis](parameter_subject_analysis.png)\n\n")

        file.write("\n#### Parameter-wise Faculty Analysis\n\n")
        file.write("| Faculty | " + " | ".join(f"Q{i}" for i in range(1, 13)) + " |\n")
        file.write("|" + "---|" * 13 + "\n")
        for faculty, param_data in analysis_result['Parameter-wise Analysis']['Faculty'].items():
            row = f"| {faculty} |"
            for param, average in param_data.items():
                row += f" {average:.2f} |"
            file.write(row + "\n")
        file.write("\n![Parameter-wise Faculty Analysis](parameter_faculty_analysis.png)\n\n")

# Usage
file_path = 'Odd_2023.csv'
output_file = 'analysis_report.md'
analysis_result, data = analyze_feedback(file_path)
generate_charts(analysis_result)
generate_markdown_report(analysis_result, data, output_file)