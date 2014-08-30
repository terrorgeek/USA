#########################################################################
#########################################################################
#
#  Extract.py
#
#  Description: extracts emails from an Excel file as a paragraph of
#  text, with each email separated by commas
#
#  Input: Any Excel file which has a column of emails. Python will find
#  this column, it need not be labeled as an email column.
#
#  Output: saves all emails into a single-paragraph file, emails.txt. The
#  emails are all separated with commas and spaces.
#
#  Procedure: Program extracts the column into an array and then writes
#  it into a text file.
#
#########################################################################
#########################################################################

#Import modules needed to run this script
import xlrd
import re

#Get name of the Excel file to extract from
filename = raw_input("Enter name of Excel file to extract e-mail addresses from: ")

#Open the Excel file
wb = xlrd.open_workbook(filename)

#Nifty email pattern to use for finding email addresses
email_pattern = re.compile('([\w\-\.]+@(\w[\w\-]+\.)+[\w\-]+)')

#Each email found in the Excel file will be appended to this string variable
#Each email in the string will be separated by a comma
email_string = ""

#Access each cell, and if it matches the email pattern, add the matched email to the email_string
#variable, followed by a comma
for sh in wb.sheets():
    for rownum in range(sh.nrows):
        row = sh.row_values(rownum)
        for cell in row:
            try: c = str(cell)
            except UnicodeEncodeError: c = cell.encode("UTF-8")
            for match in email_pattern.findall(c):
                email_string += match[0]+","

#Write email_string to extracted_emails.txt
f = open("extracted_emails.txt", "w")
f.write(email_string)
f.close()