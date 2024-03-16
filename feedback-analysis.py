import csv
from collections import defaultdict

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

            scores = [int(row[f'Q{i}']) for i in range(1, 13)]
            data[term][branch][semester][subject][faculty].append(scores)

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
                      for faculty in data[term][branch][semester][subject] for score in data[term][branch][semester][subject][faculty]]
            analysis['Branch Analysis'][branch]['Average scores for Q1-Q12'] = [calculate_average([s[i] for s in scores]) for i in range(12)]
            analysis['Branch Analysis'][branch]['Overall average'] = calculate_average([s for sub_scores in scores for s in sub_scores])

    # Semester Analysis
    analysis['Semester Analysis'] = {}
    for term in data:
        for branch in data[term]:
            for semester in data[term][branch]:
                scores = [score for subject in data[term][branch][semester] for faculty in data[term][branch][semester][subject]
                          for score in data[term][branch][semester][subject][faculty]]
                analysis['Semester Analysis'][f"Sem {semester} ({branch})"] = {'Overall average': calculate_average([s for sub_scores in scores for s in sub_scores])}

    # Subject Analysis
    analysis['Subject Analysis'] = {}
    for term in data:
        for branch in data[term]:
            for semester in data[term][branch]:
                for subject in data[term][branch][semester]:
                    if subject not in analysis['Subject Analysis']:
                        analysis['Subject Analysis'][subject] = {}
                    scores = [score for faculty in data[term][branch][semester][subject] for score in data[term][branch][semester][subject][faculty]]
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
                        scores = data[term][branch][semester][subject][faculty]
                        analysis['Faculty Analysis'][faculty]['Subjects'][subject] = calculate_average([s for sub_scores in scores for s in sub_scores])
                        if 'Overall average' not in analysis['Faculty Analysis'][faculty]:
                            analysis['Faculty Analysis'][faculty]['Overall average'] = []
                        analysis['Faculty Analysis'][faculty]['Overall average'].extend([s for sub_scores in scores for s in sub_scores])

    for faculty in analysis['Faculty Analysis']:
        analysis['Faculty Analysis'][faculty]['Overall average'] = calculate_average(analysis['Faculty Analysis'][faculty]['Overall average'])

    return analysis, data

def generate_markdown_report(analysis_result, data, output_file):
    with open(output_file, 'w') as file:
        for category, category_data in analysis_result.items():
            file.write(f"## {category}\n\n")
            if isinstance(category_data, dict):
                for item, item_data in category_data.items():
                    file.write(f"### {item}\n\n")
                    if isinstance(item_data, dict):
                        for metric, value in item_data.items():
                            if isinstance(value, list):
                                file.write(f"- {metric}: {', '.join(map(lambda x: f'{x:.2f}', value))}\n")
                            elif isinstance(value, float):
                                file.write(f"- {metric}: {value:.2f}\n")
                    else:
                        file.write(f"{item_data}\n\n")
            else:
                file.write(f"{category_data}\n\n")

        file.write("## Branch Rating Details\n\n")
        for term in data:
            for branch in data[term]:
                file.write(f"{branch}: {analysis_result['Branch Analysis'][branch]['Overall average']:.2f}\n")
                file.write(f"- {term} Term: {analysis_result['Branch Analysis'][branch]['Overall average']:.2f}\n")
                for semester in data[term][branch]:
                    file.write(f"  - Sem{semester} ({branch}): {analysis_result['Semester Analysis'][f'Sem {semester} ({branch})']['Overall average']:.2f}\n")
                    for subject in data[term][branch][semester]:
                        file.write(f"    - {subject}: {analysis_result['Subject Analysis'][subject]['Overall average']:.2f}\n")
                        for faculty in data[term][branch][semester][subject]:
                            scores = data[term][branch][semester][subject][faculty]
                            file.write(f"      - {faculty}: {calculate_average([s for sub_scores in scores for s in sub_scores]):.2f}\n")

        file.write("\n## Faculty Rating Details\n\n")
        for faculty, faculty_data in analysis_result['Faculty Analysis'].items():
            file.write(f"{faculty}: {faculty_data['Overall average']:.2f}\n")
            for subject, rating in faculty_data['Subjects'].items():
                file.write(f"- {subject}: {rating:.2f}\n")

# Usage
file_path = 'Odd_2023.csv'
output_file = 'analysis_report.md'
analysis_result, data = analyze_feedback(file_path)
generate_markdown_report(analysis_result, data, output_file)