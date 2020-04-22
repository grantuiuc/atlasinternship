# Email Data Automation Python Script Documentation ATLAS Spring 2020 

## Overview
 The email script concatenates any amount of pdfs by appending the first page of each specified pdf to one another into one large combined pdf named ‘result.pdf’. After concatenating all the report pdfs into one, it converts all of the required data into a .txt file named ‘test.txt’. By using the xlwt Python to Excel library, the program will fill in the required text onto a new Excel spreadsheet called ‘emaildata.xls’ that is already formatted for you.
 
 ## Notes
-Make sure to close AND delete 'result.pdf' and 'emaildata.xls' before running to bypass permission errors (hope to fix)
-Each run of the script will generate a new 'test.txt', 'result.pdf', and 'emaildata.xls'
-Category 1, Category 2, and Campus may need to be manually filled


 ## Install Packages/Libraries
 > pip install xlwt && pip install pypdf2 && pip install XlsxWriter && pip install python-dateutil && pip install DateTime && pip install datefinder

## Procedure
1. Create a new folder on your desktop to put our files in
2. Copy over emailToExcel.py and open it in your preferred IDE (VSCode recommended)
3. Download all email report pdfs from Webtools into the newly created folder that includes emailToExcel.py
4. Run the Python file and open up 'emaildata.xlsx' to ensure everything is filled out properly
5. Copy the filled out information to our 'Email and Social Media Data" in Box



