import csv
from collections import defaultdict

def calculate_average(scores):
    return sum(scores) / len(scores)

def analyze_feedback(file_path):
    data = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(list)))))

    with open(file_path, 'r') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            term = row['Odd/\nEven']
            branch = row['Branch']
            semester = row['Sem']
            subject = row['Subject ShortForm']
            faculty = row['Faculty Name']

            for i in range(1, 13):
                question = f'Q{i}'
                score = int(row[question])
                data[term][branch][semester][subject][faculty].append(score)

    analysis = {}

    for term, term_data in data.items():
        analysis[term] = {}
        for branch, branch_data in term_data.items():
            analysis[term][branch] = {}
            for semester, semester_data in branch_data.items():
                analysis[term][branch][semester] = {}
                for subject, subject_data in semester_data.items():
                    analysis[term][branch][semester][subject] = {}
                    for faculty, scores in subject_data.items():
                        analysis[term][branch][semester][subject][faculty] = {
                            f'Average Q{i}': calculate_average([score[i-1] for score in scores])
                            for i in range(1, 13)
                        }
                        analysis[term][branch][semester][subject][faculty]['Overall Average'] = calculate_average(
                            [score for question_scores in scores for score in question_scores]
                        )

    return analysis

# Usage
file_path = 'Feedback - Odd 2023.csv'
analysis_result = analyze_feedback(file_path)

# Print the analysis result
for term, term_data in analysis_result.items():
    print(f"Term: {term}")
    for branch, branch_data in term_data.items():
        print(f"  Branch: {branch}")
        for semester, semester_data in branch_data.items():
            print(f"    Semester: {semester}")
            for subject, subject_data in semester_data.items():
                print(f"      Subject: {subject}")
                for faculty, faculty_data in subject_data.items():
                    print(f"        Faculty: {faculty}")
                    for metric, value in faculty_data.items():
                        print(f"          {metric}: {value:.2f}")