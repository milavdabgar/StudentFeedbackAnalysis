Add Features to export as pdf as well as excel

last time you used reportlab for pdf report and openpyxl for excel report, which approach produces best reports?

Now create a flask web app that let's user download sample csv format, let them upload student feedback in csv format, performs above analysis and report generation and then offers user options to download report in markdown, excel or pdf format. Use flask with Bootstrap, wtf forms, wtf quickform, flask-sqlalchemy or any needed extension to create such web app. I believe form for file upload can be directly created using wtf-quickform.

Now in above flask app add feature to collect feedback from students, in such a way that student response are as per my earlier shared csv response file, on which we performed all student feedback analysis. In short i want to build a web app to perform student feedback collection, analysis and report generation.

Here in this approach student has to write name of each and everything, which can cause issues in data consistency. I want to have that, first student selects year, then odd/even, then branch, then semester, then subject then faculty and then can rate this combination on the scale of 1-5 for Q1 to Q12. All of the options should be either populated from database of predefined options only, nothing to be typed. Please use this approach.


For Branch, Semester, Subject Code, Name, Faculty Name and Faculty Subjects, Instead of dummy values, use real values from shared csv file.