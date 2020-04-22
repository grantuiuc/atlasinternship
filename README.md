# Email Data Automation Python Script Documentation ATLAS Spring 2020 

## Overview
   The email script is created to automate and efficiently code SMMC's email data instead of manually filling out each field. It can help save lots of time and tedious work by filling out any number of reports onto an Excel file :smiley:. It is primarily dynamic as it will continue to work with all of our Webtool's pdf reports but can fall to some unaccounted issues (Specified below) :upside_down_face:. Hopefully this program will continue to be improved upon and finalized in the near future! :sunglasses:
I have included comments to describe what the program does line-by-line for easier understanding and future improvements.

 
## How it works
   The python script dynamically searches for key words in the pdfs to fill our required information on the Excel sheet. It utilizes python libraries that allow the program to interact with Excel sheets and real-time dates. The program concatenates any amount of pdfs by appending the first page of each specified pdf to one another into one large combined pdf named ‘result.pdf’. After concatenating all the report pdfs into one, it converts all of the required data into a .txt file named ‘test.txt’. By using the xlwt Python to Excel library, the program will fill in the required text onto a new Excel spreadsheet called ‘emaildata.xls’ that is already formatted for you. Certain keywords may contain multiple duplicates but it accounts for this by only considering every other keyword in those situations. 
 
 ## Notes
- __Make sure to close *AND* delete 'result.pdf' and 'emaildata.xls' before running to bypass permission errors (hope to fix)__
- Each run of the script will generate a new 'test.txt', 'result.pdf', and 'emaildata.xls'
- Please open pdfs in your own local machine and not on your local IDE
- When filling out the Subject, it looks for the specific term 'SMMC'. However, not every email subject contains this keyword. 
- Category 1, Category 2, and Campus may need to be manually filled


 ## Install Packages/Libraries
 > pip install xlwt && pip install pypdf2 && pip install XlsxWriter && pip install python-dateutil && pip install DateTime && pip install datefinder

## Procedure
1. Create a new folder on your desktop to put our files in
2. Copy over emailToExcel.py and open it in your preferred IDE (VSCode recommended)
3. Download all email report pdfs from Webtools into the newly created folder that includes emailToExcel.py
4. Run the Python file and open up 'emaildata.xlsx' to ensure everything is filled out properly
5. Copy the filled out information to our master 'Email and Social Media Data" file in Box



