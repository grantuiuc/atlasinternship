#@author Grant Chiu grantc2

import PyPDF2 
import sys
import os
from PyPDF2 import PdfFileMerger
import pandas as pd
import xlwt 
from xlwt import Workbook 
import xlsxwriter
import re
from datetime import datetime
import datefinder
import datetime
import calendar
from dateutil import parser
#MAKE SURE TO CLOSE result.pdf/emaildata.xls BEFORE RUNNING SCRIPT OR ELSE IT WILL GIVE OUT A 'PERMISSION DENIED ERROR'

#name of the pdfs you want to append
pdfs = []

for filename in os.listdir('.'):
    if filename.endswith('.pdf'):
        pdfs.append(filename)

for file in pdfs:
    print(file)

merger = PdfFileMerger()

for pdf in pdfs:
    merger.append(pdf, pages=(0,1))

#appends all pdfs to a pdf called result.pdf, please open this in a pdf reader and not your local IDE
merger.write("result.pdf")
merger.close()


#Creates a pdf file object
pdfFileObj = open('result.pdf', 'r+b') 
  
#Creates a pdf reader object
pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
  
#Prints number of pages in the pdf, may be unecessary
print("Page Count: "+ str(pdfReader.numPages))
sys.stdout = open('test.txt', 'w')


#loops through all the apges
for i in range(0, pdfReader.numPages):
    #Reading everything from first page, indexes start at 0
    pageObj = pdfReader.getPage(i) 
    
    # extracting text from page 
    print(pageObj.extractText()) 
  
# closing the pdf file object 
pdfFileObj.close() 

#reset stdout
sys.stdout = sys.__stdout__

workbook = xlsxwriter.Workbook('emaildata.xlsx')
sheet1 = workbook.add_worksheet()

#Change font of text on sheet
def style():
    style = xlwt.XFStyle() 
    font = xlwt.Font()  
    font.name = 'Times New Roman'  
    font.height = 12
    style = xlwt.XFStyle() # Create the Style
    style.font = font 
    return style

#creates headers for excel sheet
def create_headers():
    row = 0
    column = 0
    content = ["Sent Date", "Subject Line", "Category 1", "Category 2", 
                    "Campus", "Type", "Total Sent", "Total Opens","Unique Opens","Not Opened", "Bounced", 
                    "Number of Clicks","Desktop Opens","Mobile Opens","Day of Week", "Open %","Click %","Desktop %","Mobile %"] 

    # iterating through content list 
    for item in content: 
        # write operation perform 
        sheet1.write(row, column, item) 
        column += 1
create_headers()



#counts file length and stores it into a variable called number_of_lines
def file_len(fname):
    num_lines = sum(1 for line in open(fname))
    return num_lines
number_of_lines = file_len('test.txt')

#read from test.txt to string
lines = []
with open('test.txt') as f:
    lines = f.readlines()
lines = [x.strip() for x in lines] 


#Fill all of the columns
def fill_sent_date():
    row = 1
    col = 0
    row1 = 1
    col2 = 14
    term = "Sent Date"
    lines = []
    day_of_week = []
    file = open('test.txt')
    for line in file:
        line.strip().split('/n')
        if term in line:
            print(line)
            print(parser.parse(line[11:-9]).strftime("%A"))
            day = parser.parse(line[11:-9]).strftime("%A")
            lines.append(line[11:-9])
            day_of_week.append(day)
    for i in range(len(lines)):
        sheet1.write(row, col, lines[i])
        row += 1
    for i in range(len(day_of_week)):
        sheet1.write(row1, col2, day_of_week[i])
        row1 += 1
    file.close()


def fill_subjects():
    file = open('test.txt')
    row = 1
    col = 1
    term = "SMMC"
    lines = []
    file = open('test.txt')
    for line in file:
        line.strip().split('/n')
        if term in line:
            lines.append(line)
    for i in range(len(lines)):
        sheet1.write(row, col, lines[i])
        row += 1
    file.close()

def fill_total_sent():
    file = open('test.txt')
    row = 1
    col = 6
    term = "Sent\n"
    lines = []
    file = open('test.txt')
    current_line = file.readline()
    for line in file:
        line.strip().split('/n')
        previous_line = current_line
        current_line = line
        if term in line:
            lines.append(previous_line)
    for i in range(len(lines)):
        sheet1.write(row, col, lines[i])
        row += 1
    file.close()

def fill_total_opens():
    file = open('test.txt')
    row = 1
    col = 7
    term = "Opens\n"
    lines = []
    file = open('test.txt')
    current_line = file.readline()
    every_other = 0

    for line in file:
        line.strip().split('/n')
        previous_line = current_line
        current_line = line
        if term in line:
            lines.append(previous_line)
    for i in range(len(lines)):
        if every_other % 2 == 0:
            sheet1.write(row, col, lines[i])
            row += 1
        every_other += 1
    file.close()

def fill_unique_opens():
    file = open('test.txt')
    row = 1
    col = 8
    term = "Unique full opens"
    lines = []
    file = open('test.txt')
    current_line = file.readline()

    for line in file:
        line.strip().split('/n')
        previous_line = current_line
        current_line = line
        if term in line:
            lines.append(previous_line)
    for i in range(len(lines)):
        sheet1.write(row, col, lines[i])
        row += 1
    file.close()

def fill_not_opened():
    file = open('test.txt')
    row = 1
    col = 9
    term = "Not opened"
    lines = []
    file = open('test.txt')
    current_line = file.readline()

    for line in file:
        line.strip().split('/n')
        previous_line = current_line
        current_line = line
        if term in line:
            lines.append(previous_line)
    for i in range(len(lines)):
        sheet1.write(row, col, lines[i])
        row += 1
    file.close()

def fill_bounced():
    file = open('test.txt')
    row = 1
    col = 10
    term = "Bounced"
    lines = []
    file = open('test.txt')
    current_line = file.readline()

    for line in file:
        line.strip().split('/n')
        previous_line = current_line
        current_line = line
        if term in line:
            lines.append(previous_line)
    for i in range(len(lines)):
        sheet1.write(row, col, lines[i])
        row += 1
    file.close()

def fill_number_of_clicks():
    file = open('test.txt')
    row = 1
    col = 11
    term = "People clicked"
    lines = []
    file = open('test.txt')
    current_line = file.readline()

    for line in file:
        line.strip().split('/n')
        previous_line = current_line
        current_line = line
        if term in line:
            lines.append(previous_line)
    for i in range(len(lines)):
        sheet1.write(row, col, lines[i])
        row += 1
    file.close()

def fill_desktop_opens():
    file = open('test.txt')
    row = 1
    col = 12
    term = "Desktop opens"
    lines = []
    file = open('test.txt')
    current_line = file.readline()

    for line in file:
        line.strip().split('/n')
        previous_line = current_line
        current_line = line
        if term in line:
            lines.append(previous_line)
    for i in range(len(lines)):
        sheet1.write(row, col, lines[i])
        row += 1
    file.close()

def fill_mobile_opens():
    file = open('test.txt')
    row = 1
    col = 13
    term = "Mobile opens"
    lines = []
    current_line = file.readline()

    for line in file:
        line.strip().split('/n')
        previous_line = current_line
        current_line = line
        if term in line:
            lines.append(previous_line)
    for i in range(len(lines)):
        sheet1.write(row, col, lines[i])
        row += 1
    file.close()

def fill_open_rate():
    file = open('test.txt')
    row = 1
    col = 15
    term = "open rate"
    lines = []
    for line in file:
        line.strip().split('/n')
        if term in line:
            line_as_number = re.findall('\d+\.\d+%', line)
            lines.append(line_as_number[0])
    for i in range(len(lines)):
        sheet1.write(row, col, lines[i])
        row += 1
    file.close()

def fill_click_rate():
    file = open('test.txt')
    row = 1
    col = 16
    term = "click rate"
    lines = []
    for line in file:
        line.strip().split('/n')
        if term in line:
            line_as_number = re.findall('\d+\.\d+%', line)
            lines.append(line_as_number[0])
    for i in range(len(lines)):
        sheet1.write(row, col, lines[i])
        row += 1
    file.close()

fill_subjects()
fill_total_sent()
fill_total_opens()
fill_unique_opens()
fill_not_opened()
fill_bounced()
fill_number_of_clicks()
fill_desktop_opens()
fill_mobile_opens()
fill_sent_date()
fill_open_rate()
fill_click_rate()

rowcount=2
# sheet1.write(16, 1, xlwt.Formula("=NUMBERVALUE(H2) / NUMBERVALUE(G2)"))
# sheet1.write(10, 10, sheet1.write_formula('P2', '=_xlfn.NUMBERVALUE(h2) / _xlfn.NUMBERVALUE(g2) '))
# formula = sheet1.write_formula('P' + str(rowcount), '=_xlfn.NUMBERVALUE(h2) / _xlfn.NUMBERVALUE(g2) ')


workbook.close()