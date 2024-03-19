Add Features to export as pdf as well as excel

last time you used reportlab for pdf report and openpyxl for excel report, which approach produces best reports?

Now create a flask web app that let's user download sample csv format, let them upload student feedback in csv format, performs above analysis and report generation and then offers user options to download report in markdown, excel or pdf format. Use flask with Bootstrap, wtf forms, wtf quickform, flask-sqlalchemy or any needed extension to create such web app. I believe form for file upload can be directly created using wtf-quickform.

Now in above flask app add feature to collect feedback from students, in such a way that student response are as per my earlier shared csv response file, on which we performed all student feedback analysis. In short i want to build a web app to perform student feedback collection, analysis and report generation.

Here in this approach student has to write name of each and everything, which can cause issues in data consistency. I want to have that, first student selects year, then odd/even, then branch, then semester, then subject then faculty and then can rate this combination on the scale of 1-5 for Q1 to Q12. All of the options should be either populated from database of predefined options only, nothing to be typed. Please use this approach.


For Branch, Semester, Subject Code, Name, Faculty Name and Faculty Subjects, Instead of dummy values, use real values from shared csv file.



## Feedback Analysis

## Branch Analysis (overall)

| Branch | Average Score |
|--------|---------------|
| EC | 4.56 |
| ICT | 3.73 |
| IT | 4.19 |

## Semester Analysis (overall)

| Branch - Semester | Average Score |
|----------|---------------|
| EC - Sem 5 | 4.99 |
| EC - Sem 3 | 4.26 |
| EC - Sem 1 | 4.35 |
| ICT - Sem 3 | 4.30 |
| ICT - Sem 1 | 3.56 |
| IT - Sem 1 | 4.19 |

### Subject Analysis (overall)

| Subject | Overall Average |
|---------|------------------|
| E&S (4300021) | 4.97 |
| ES (4351102) | 5.00 |
| MWR (4351103) | 4.98 |
| M&WC (4351104) | 5.00 |
| SP (4351105) | 5.00 |
| OPP (4351108) | 5.00 |
| PR1 (4351107) | 5.00 |
| ECN (4331101) | 4.52 |
| EMI (4331102) | 3.78 |
| IE (4331103) | 4.19 |
| PEC (4331104) | 4.61 |
| PC (4331105) | 4.19 |
| FE (4311102) | 4.92 |
| FEE (4311101) | 4.94 |
| BICT (4300010) | 3.93 |
| S&Y (4300015) | 4.29 |
| CE (1333201) | 4.09 |
| MPMC (1333202) | 4.15 |
| DSA (1333203) | 4.33 |
| DBMS (1333204) | 4.35 |
| OSA (1333205) | 4.58 |
| EEE (1313202) | 3.94 |
| WDP (1313203) | 3.64 |
| FICT (1313201) | 3.40 |
| PP (4311601) | 4.43 |
| SWD (4311603) | 4.36 |
| IIS (4311602) | 4.02 |

### Faculty Analysis (Overall)

| Subject | Overall Average |
|---------|------------------|
| Mr. S P Joshiara | 4.27 |
| Mr. M J Vadhwania | 3.56 |


## Faculty Analysis (Subject-wise)

### Mr. S P Joshiara

- Overall Average: 4.27

| Subject | Average Score |
|---------|---------------|
| E&S (4300021) | 5.00 |
| EMI (4331102) | 3.25 |

### Mr. M J Vadhwania

- Overall Average: 3.56

| Subject | Average Score |
|---------|---------------|
| BICT (4300010) | 1.58 |
| S&Y (4300015) | 4.08 |
| IIS (4311602) | 5.00 |


## Subject Analysis (Faculty-wise)

### S&Y (4300015)

- Overall Average: 4.39

| Faculty | Average Score |
|---------|---------------|
| Mr. M J Vadhwania | 4.08 |
| Mr. N J Chauhan | 4.08 |
| Mr. S P Joshiara | 5.00 |

### CE (1333201)

- Overall Average: 3.83

| Faculty | Average Score |
|---------|---------------|
| Mr. S J Chauhan | 3.83 |


#### Branch Analysis (Parameter-wise)

| Branch | Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 | Overall |
|----------|----|----|----|----|----|----|----|----|----|-----|-----|-----|-------| 4.58 | 
| EC | 4.63 | 4.44 | 4.48 | 4.52 | 4.53 | 4.64 | 4.59 | 4.60 | 4.54 | 4.49 | 4.59 | 4.64 | 4.29 |
| ICT | 3.73 | 3.63 | 3.74 | 3.76 | 3.72 | 3.74 | 3.75 | 3.72 | 3.73 | 3.71 | 3.73 | 3.78 | 4.5 |
| IT | 4.26 | 4.16 | 4.21 | 4.20 | 4.12 | 4.31 | 4.23 | 4.21 | 4.18 | 4.04 | 4.19 | 4.17 | 4.13 |

#### Subject Analysis (Parameter-wise)

| Subject | Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |
|----------|----|----|----|----|----|----|----|----|----|-----|-----|-----|
| E&S (4300021) | 5.00 | 5.00 | 5.00 | 5.00 | 4.80 | 5.00 | 5.00 | 4.80 | 5.00 | 5.00 | 5.00 | 5.00 |  4.05 |
| ES (4351102) | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 |  4.06 |
| MWR (4351103) | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 4.80 |  4.80 |
| M&WC (4351104) | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 |  4.20 |
| SP (4351105) | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 5.00 | 4.20 |

#### Faculty Analysis (Parameter-wise)

| Faculty | Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |
|----------|----|----|----|----|----|----|----|----|----|-----|-----|-----|
| Mr. S P Joshiara | 3.59 | 3.48 | 3.54 | 3.51 | 3.49 | 3.58 | 3.58 | 3.55 | 3.52 | 3.50 | 3.56 | 3.57 |
| Ms. M K Pedhadiya | 4.07 | 3.85 | 4.11 | 4.10 | 4.00 | 4.08 | 4.05 | 4.08 | 4.08 | 4.18 | 4.20 | 4.21 |
| Mr. R C Parmar | 3.93 | 3.94 | 4.00 | 3.99 | 3.90 | 4.11 | 3.93 | 4.00 | 3.91 | 3.79 | 3.97 | 3.97 |
| Mr. L K Patel | 4.56 | 4.64 | 4.74 | 4.78 | 4.68 | 4.72 | 4.68 | 4.70 | 4.78 | 4.68 | 4.72 | 4.82 |
| Mr. R N Patel | 4.01 | 3.93 | 3.93 | 4.00 | 3.94 | 4.07 | 3.99 | 3.90 | 4.00 | 3.87 | 3.99 | 3.98 |
| Mr. M J Dabgar | 4.50 | 4.31 | 4.50 | 4.48 | 4.50 | 4.45 | 4.52 | 4.43 | 4.52 | 4.26 | 4.55 | 4.55 |
| Mr. S J Chauhan | 4.32 | 3.77 | 4.23 | 4.41 | 4.41 | 4.32 | 4.36 | 4.32 | 4.18 | 4.27 | 4.32 | 4.45 |
| Mr. N J Chauhan | 4.08 | 4.04 | 3.99 | 4.07 | 4.08 | 4.09 | 4.19 | 4.15 | 4.09 | 4.08 | 4.00 | 4.08 |
| Mr. M J Vadhwania | 3.66 | 3.28 | 3.21 | 3.14 | 3.10 | 3.24 | 3.21 | 3.21 | 2.93 | 3.03 | 2.86 | 2.90 |


## Semester Analysis (Parameter-wise)

| Branch - Semester | Average Score |
|----------|---------------|
| EC - Sem 5 | 4.99 |
| EC - Sem 3 | 4.26 |
| EC - Sem 1 | 4.35 |
| ICT - Sem 3 | 4.30 |
| ICT - Sem 1 | 3.56 |
| IT - Sem 1 | 4.19 |

For my flask webapp that takes feedback.csv as input, and generated feedback analysis report as output. 

Create analyze_feedback(file_content) function that takes the content of a feedback file as input which can be used in below flask route

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

create generate_markdown_report(analysis_result, markdown_file) function which can be used in below function. 

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

I Want markdown report to have tables as specified below. 
## Feedback Analysis

## Overall Feedback Analysis

| Branch Score | Term-Year Score |  Semester Score | Subject Score |

### Branch Analysis (overall)

| Branch | Average Score |
|--------|---------------|

### Term-Year Analysis (overall)

| Term-Year | Overall |

### Semester Analysis (overall)

| Branch - Semester | Average Score |
|----------|---------------|

### Subject Analysis (overall)

| Subject | Overall Average |
|---------|------------------|

### Faculty Analysis (Overall)

| Subject | Overall Average |
|---------|------------------|

## Parameter-wise Feedback Analysis

### Branch Analysis (Parameter-wise)

| Branch | Overall |Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |

### Term-Year Analysis (Parameter-wise)

| Term-Year | Overall |Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |

### Semester Analysis (Parameter-wise)

| Branch - Semester | Overall |Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |

### Subject Analysis (Parameter-wise)

| Subject | Overall |Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |

### Faculty Analysis (Parameter-wise)

| Faculty | Overall |Q1 | Q2 | Q3 | Q4 | Q5 | Q6 | Q7 | Q8 | Q9 | Q10 | Q11 | Q12 |

## Misc Feedback Analysis

### Subject Analysis (Faculty-wise)

#### S&Y (4300015)

- Overall Average: 4.39

| Faculty | Average Score |
|---------|---------------|
| Mr. M J Vadhwania | 4.08 |
| Mr. N J Chauhan | 4.08 |
| Mr. S P Joshiara | 5.00 |

#### CE (1333201)

- Overall Average: 3.83

| Faculty | Average Score |
|---------|---------------|
| Mr. S J Chauhan | 3.83 |

### Faculty Analysis (Subject-wise)

#### Mr. S P Joshiara

- Overall Average: 4.27

| Subject | Average Score |
|---------|---------------|
| E&S (4300021) | 5.00 |
| EMI (4331102) | 3.25 |

#### Mr. M J Vadhwania

- Overall Average: 3.56

| Subject | Average Score |
|---------|---------------|
| BICT (4300010) | 1.58 |
| S&Y (4300015) | 4.08 |
| IIS (4311602) | 5.00 |

Use bottom up approach to calculate scores. 

First create subject scores (faculty-wise), use it to create subject score (overall.) generated subject score should tally with average of subject score (faculty-wise). 
Then calculate faculty scores (subject-wise), use it to create faculty score (overall.) generated faculty score should tally with average of faculty score (subject-wise). 

then use subject score (overall) to calculate semester score. semester score should tally with average of subjects of a particular semester. 

then use semester score to calculate branch score. branch score should tally with average of semesters.

then use branch score to calculate term-year (Odd-2023). term-year score should tally with average of semesters of a particular branch.

for calculating any of the scores above, make sure that you create sample having earlier values unique.

e.g. while calculating semester scores. make sure that data[year_term][branch] is unique. dont club semester 3 of EC Branch with semester 3 of ICT Branch. 



aren't the problem same for Faculty Analysis (Parameter-wise) & Subject Analysis (Parameter-wise) but for one you suggested change is analyze\\\_feedback() methods and for other in generate\\\_markdown\\\_report(). is there any particular reason?