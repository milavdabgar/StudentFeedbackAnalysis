Add Features to export as pdf as well as excel

last time you used reportlab for pdf report and openpyxl for excel report, which approach produces best reports?

Now create a flask web app that let's user download sample csv format, let them upload student feedback in csv format, performs above analysis and report generation and then offers user options to download report in markdown, excel or pdf format. Use flask with Bootstrap, wtf forms, wtf quickform, flask-sqlalchemy or any needed extension to create such web app. I believe form for file upload can be directly created using wtf-quickform.

Now in above flask app add feature to collect feedback from students, in such a way that student response are as per my earlier shared csv response file, on which we performed all student feedback analysis. In short i want to build a web app to perform student feedback collection, analysis and report generation.

Here in this approach student has to write name of each and everything, which can cause issues in data consistency. I want to have that, first student selects year, then odd/even, then branch, then semester, then subject then faculty and then can rate this combination on the scale of 1-5 for Q1 to Q12. All of the options should be either populated from database of predefined options only, nothing to be typed. Please use this approach.


For Branch, Semester, Subject Code, Name, Faculty Name and Faculty Subjects, Instead of dummy values, use real values from shared csv file.



## Feedback Analysis

## Faculty Analysis

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

## Subject Analysis

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

## Semester Analysis

| Branch - Semester | Average Score |
|----------|---------------|
| EC - Sem 5 | 4.99 |
| EC - Sem 3 | 4.26 |
| EC - Sem 1 | 4.35 |
| ICT - Sem 3 | 4.30 |
| ICT - Sem 1 | 3.56 |
| IT - Sem 1 | 4.19 |

## Branch Analysis

| Branch | Average Score |
|--------|---------------|
| EC | 4.56 |
| ICT | 3.73 |
| IT | 4.19 |