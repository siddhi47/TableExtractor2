# TableExtractor2
A table extractor that extracts tables from pdf files (after matching a string). 


<b><========================<SECTION A - Prologue>========================></b>
<br>
IF YOU ARE NOT INTERESTED IN HOW I FINALLY WROTE THIS PROGRAM. YOU MAY SKIP TO SECTION B. BUT I HIGHLY RECOMMEND THAT YOU DO READ! (The more you read the more happy i become)

USING PYTHON LIBRARIES FOR CONVERSION
Many hours were wasted trying to convert pdf into excel/csv using tabula (which allegedly is one of the most advanced libraries used to do the same)

I also tried my hands on PyPDF2. but in vain, produced results that were not consistent.


USING THIRD PARTY SOFTWARE FOR CONVERSION (WORKS ONLY IN WINDOWS)
Then I thought of using our old trusty office package to achieve what these libraries could not achieve. The results were obviously better 
than these libraries. The main idea was to open the pdf in word application, save it in an html format, then scrape the table from the 
html file. This was too tedious to do manually. So i used pywinauto to achive this automation. After many attemts trying to figure out
how pywinauto worked. I finally automated these tasks. (So, if you want to learn about pywinauto, don't ask me! :( )

Even After automating the tasks for office package, I was still left unsatisfied. Then I turned to use adobe acrobat dc. It had an option to
save pdf as excel spreadsheet. Finally I managed to automate the whole process (using same pywinauto :( )


<b><========================<SECTION B - The Real Read Me>========================></b>
The following should be installed in your system:
	1. pywinauto (https://pywinauto.github.io/)
	2. adobe acrobat dc

Make slight changes in your code:
Look at the comments at the line 60 and 78

Run the following command: 
python tableextractor.py -l [List of urls] -k [keyword to match]

Example
python tableextractor.py -l 'https://www.nibl.com.np/images/AnnualReport/Final%20Financial%20Statements%2073-74.pdf' ,'https://www.nabilbank.com/images/pdf/uploaded/Annual%20Report_16-17.pdf' -k 'House Rent'
