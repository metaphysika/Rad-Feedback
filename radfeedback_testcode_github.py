# -*- coding: utf-8 -*-
"""
Created on 1-12-18

@author: clahn
"""

import os, sys
import win32com.client as win32
import py
import pandas as pd
import shutil
import time
from emailsender import *

# this checks if outlook is open.  If not, opens it.
EmailSender().check_outlook()

# snippet to become date in file name below.
todaydate2 = time.strftime("%m%d%Y")
# naming of daily report for archival. The "/" was necessary to set the file path along with the todaydate function.
dailyreportname = ("W:\\SHARE8 Physics\\Software\\python\\scripts\\clahn\\EPIC Rad Feedback\\incoming_epic_rad_reports/" +
                    todaydate2 + " daily_radfeedback_equip_report" + ".xlsx")
dailyreportname2 = ("W:\\SHARE8 Physics\\Software\\python\\scripts\\clahn\\EPIC Rad Feedback\\incoming_epic_rad_reports/" +
                    todaydate2 + " daily_radfeedback_equip_report" + ".csv")


# This will unprotect workbook and save it again.
xcl = win32.Dispatch('Excel.Application')
pw_str = ''
wb = xcl.Workbooks.Open(dailyreportname, False, False, None, pw_str)
xcl.DisplayAlerts = False
wb.SaveAs(dailyreportname, None, '', '')
xcl.Quit()


# this converts excel file to csv file.
df = pd.read_excel(dailyreportname).to_csv(dailyreportname2, index=False)


# Open daily report, save, and close.  Excel formats the data slightly after a save.
# This gets the daily reprot into the exact
# same format as the Master report so the duplicates are truly duplicates and get removed.
# Excel removes double quotes around data.
todaydate2 = time.strftime("%m%d%Y")
# naming of daily report for archival. The "/" was necessary to set the file path along with the todaydate function.
dailyreportname2 = ("W:\\SHARE8 Physics\\Software\\python\\scripts\\clahn\\EPIC Rad Feedback\\incoming_epic_rad_reports/" +
                    todaydate2 + " daily_radfeedback_equip_report" + ".csv")
xl = win32.DispatchEx("Excel.Application")
wb = xl.workbooks.open(dailyreportname2)
xl.Visible = False
xl.DisplayAlerts = False
wb.Save()
wb.Close()
xl.Quit()


# Load master data csv
fileMaster = py.path.local(r'W:\SHARE8 Physics\Software\python\scripts\clahn\EPIC Rad Feedback\master_report\masterreport.csv')
if not fileMaster.isfile():
    raise ValueError()
try:
    dfMaster = pd.read_csv(fileMaster, usecols=['MRN', 'Accession #', 'Begin Exam Time', 'Quality User',
                                                'Quality Element', 'Quality Comment', 'Technologist',
                                                'Dept', 'Category', 'Procedure'])

    dfMaster["MRN"] = dfMaster["MRN"].astype(str)
    dfMaster["Accession #"] = dfMaster["Accession #"].astype(str)
    dfMaster["Begin Exam Time"] = dfMaster["Begin Exam Time"].astype(str)
    dfMaster["Quality User"] = dfMaster["Quality User"].astype(str)
    dfMaster["Quality Element"] = dfMaster["Quality Element"].astype(str)
    dfMaster["Quality Comment"] = dfMaster["Quality Comment"].astype(str)
    dfMaster["Technologist"] = dfMaster["Technologist"].astype(str)
    dfMaster["Dept"] = dfMaster["Dept"].astype(str)
    dfMaster["Category"] = dfMaster["Category"].astype(str)
    dfMaster["Procedure"] = dfMaster["Procedure"].astype(str)
except Exception:
    pass



# Loop through each file report, loading CSV data, and appending that data to the master data frame.
dirname = py.path.local(r'W:\SHARE8 Physics\Software\python\scripts\clahn\EPIC Rad Feedback\incoming_epic_rad_reports')
files = []
# for f in dirname.visit(fil='*.xlsx', bf=True):
for f in dirname.visit(fil='*.csv', bf=True):
    files.append(f)
if (len(files) > 1):
    raise ValueError()
elif not len(files):
    # raise ValueError()
    # If there is no file in the incoming folder, this will email a message saying no file in folder.
    EmailSender().send_email("christopher.lahn@sanfordhealth.org", "No EPIC Master Report",
                             "There was not a file in the incoming_epic_rad_reports folder.")
files = files[0]

df = pd.read_csv(files, usecols=['MRN', 'Accession #', 'Begin Exam Time', 'Quality User',
                                 'Quality Element', 'Quality Comment', 'Technologist',
                                 'Dept', 'Category', 'Procedure'])

# print (df)


df["MRN"] = df["MRN"].astype(str)
df["Accession #"] = df["Accession #"].astype(str)
df["Begin Exam Time"] = df["Begin Exam Time"].astype(str)
df["Quality User"] = df["Quality User"].astype(str)
df["Quality Element"] = df["Quality Element"].astype(str)
df["Quality Comment"] = df["Quality Comment"].astype(str)
df["Technologist"] = df["Technologist"].astype(str)
df["Dept"] = df["Dept"].astype(str)
df["Category"] = df["Category"].astype(str)
df["Procedure"] = df["Procedure"].astype(str)

# This adds an index of master and daily to each respective dataframe.
# set_index allows the drop duplicates to work despite the order they appear in either database.
dfMaster['master'] = 'master'
dfMaster.set_index('master', append=True, inplace=True)
df['daily'] = 'daily'
df.set_index('daily', append=True, inplace=True)


# This merges the dataframes and then drops the duplicates
merged = dfMaster.append(df)
merged = merged.drop_duplicates().sort_index()
# print(merged)

# creates a new daily dataframe slicing the merged dataframe by just "daily" index.
idx = pd.IndexSlice
df = merged.loc[idx[:, 'daily'], :]

# Append the filtered daily report to the dfMaster
dfMaster = dfMaster.append(df, ignore_index=True)

# store the master data frame.  Does not include index.
dfMaster.to_csv(fileMaster, index=False)

# Write daily report dataframe to csv for Archive and email to supervisors.
# Create current date for file name of archived daily report
todaydate = time.strftime("%Y%m%d")
# naming of daily report for archival. The "/" was necessary to set the file path along with the todaydate function.
dailyreportname3 = "W:\\SHARE8 Physics\\Software\\python\\scripts\\clahn\\EPIC Rad Feedback\\archived_daily_reports/" + \
    todaydate + " daily_radfeedback_equip_processed_report" + ".csv"

# save daily report to archival folder.  Does not include index.
df.to_csv(dailyreportname3, index=False)


try:
    fileMaster.copy(py.path.local(r'\\internal.sanfordhealth.org@SSL\DavWWWRoot\departments'
                                  r'\radiology\private\fargo\qcforms\Shared Documents\EPIC_Rad_Feedback_Master'))
    fileMaster.copy(py.path.local(r'\\internal.sanfordhealth.org@SSL\DavWWWRoot\departments'
                                  r'\radiology\private\RadBIS\physics\Shared Documents\epic_rad_feedback_master'))
    fileMaster.copy(py.path.local(r'\\internal.sanfordhealth.org@SSL\DavWWWRoot\departments'
                                  r'\radiology\bemidji\Shared Documents\EPIC_Rad_Feedback_Master'))
except Exception as err:
    err = str(err)
    # This emails  any error reports to me
    # You have to be signed into to SharePoint for this to work.  If you leave your computer runnning it will stay signed in until you logout of your computer or it restarts.  You simply can sign back in by trying to "Open with Explorer" and using sanford login on pop up window.  Then manually trigger program again.
    EmailSender().send_email("christopher.lahn@sanfordhealth.org",
                             "Check EPIC Master Report: Exception Raised",
                             "You may not have been sigend into Sharepoint.  "
                             "An exception was raised: " + err)
    pass


# This sends the daily report emails to supervisors
EmailSender().send_email(, dailyreportname3)


# create dictionary for correcting email address.  Useful for people who have same name or names sith St. prefix.
d = {}


# This creates a mask of just the reasons we want.  It then itterates through the masked data frame and triggers the email for each row.
for idx, row in df.iterrows():
    # Ignore rows that have blank Quality Element
    if row.at["Quality Element"] == "" or row.at["Quality Element"] == " ":
        emailname = "christopher.lahn@sanfordhealth.org"
    # Iterate over dataframe and check "d" dictionary for problem email addresses.
    elif row.at["Technologist"] in d.keys():
        emailname = d.get(row.at["Technologist"])
    else:
        getname = row.at["Technologist"]
        first = getname.split(",")[1].split(" ")[1]
        last = getname.split(",")[0]
        emailname = first + "." + last + "@sanfordhealth.org"
    getacc = row.at["Accession #"]
    getdate = row.at["Begin Exam Time"]
    strdate = str(getdate)
    stracc = str(getacc)
    getrad = row.at['Quality User']
    getproc = row.at["Procedure"]
    getreason = row.at["Quality Element"]
    getnotes = row.at["Quality Comment"]

    # send_notification()
    EmailSender().send_email(emailname, "Automated Message:  Image Quality Feedback",
                             "Hello, \r\n \r\nThis is an automated message.  No reply is necessary."
                             "  \r\n \r\nAn image that you completed in Radiant was flagged for image quality"
                             " review by a radiologist.  \r\n\r\nPlease, use the accession number provided here"
                             " to look up your exam in PACS and review the image along with the radiologist's feedback"
                             " provided in this email.  \r\n\r\nNOTES: \r\n1. You may need to add a zero at the beginning"
                             " of the accession number to find it in PACS. \r\n2. Please, respond via email to the"
                             " radiologist listed below if they requested so in the Radiologist Notes section."
                             "\r\n3. nan = no value entered by Radiologist \r\n4. If Radiologist Reason for Image Flag"
                             " includes or is only Physics, the exam was flagged for possible equipment issues and will"
                             " be reviewed by Imaging Physics, as well.  Contact Imaging Physics if you have any"
                             " input which may help us resolve the potential equipment issue.\r\n \r\n"
                             "       Accession #: " + stracc + "\r\n \r\n       Exam Date: " + strdate + "\r\n \r\n"
                             "       Procedure: " + getproc + "\r\n \r\n        Radiologist: " + getrad + "\r\n \r\n"
                             "       Radiologist Reason For Image Flag: "
                             + getreason + "\r\n \r\n       Radiologist Notes (not always included by radiologist): "
                             + getnotes + "\r\n \r\nPlease, contact your supervisor if you have any further questions."
                             "  \r\n \r\nIf you received this message by mistake, contact physics@sanfordhealth.org."
                             " \r\n \r\nThank you, \r\n \r\nImaging Physics \r\nphysics@sanfordhealth.org")

# emails physics@sanfordhealth.org with just the physics flagged items
for idx, row in df[df["Quality Element"].str.contains("Physics")].iterrows():
    emailname = "physics@sanfordhealth.org; christopher.lahn@sanfordhealth.org"
    getacc = row.at["Accession #"]
    stracc = str(getacc)
    getdate = row.at["Begin Exam Time"]
    strdate = str(getdate)
    gettech = row.at['Technologist']
    getrad = row.at['Quality User']
    getproc = row.at["Procedure"]
    getreason = row.at["Quality Element"]
    getnotes = row.at["Quality Comment"]

    # send_notification()
    EmailSender().send_email(emailname, "Automated Message:  Image Quality Feedback",
                             "Hello, \r\n \r\nThis is an automated message.  No reply is necessary."
                             "  \r\n \r\nAn image that you completed in Radiant was flagged for image quality"
                             " review by a radiologist.  \r\n\r\nPlease, use the accession number provided here"
                             " to look up your exam in PACS and review the image along with the radiologist's feedback"
                             " provided in this email.  \r\n\r\nNOTES: \r\n1. You may need to add a zero at the beginning"
                             " of the accession number to find it in PACS. \r\n2. Please, respond via email to the"
                             " radiologist listed below if they requested so in the Radiologist Notes section."
                             "\r\n3. nan = no value entered by Radiologist \r\n4. If Radiologist Reason for Image Flag"
                             " includes or is only Physics, the exam was flagged for possible equipment issues and will"
                             " be reviewed by Imaging Physics, as well.  Contact Imaging Physics if you have any"
                             " input which may help us resolve the potential equipment issue.\r\n \r\n"
                             "       Accession #: " + stracc + "\r\n \r\n       Exam Date: " + strdate +
                             "\r\n \r\n       Procedure: " + getproc + "\r\n \r\n       Tech: " + gettech +
                             "\r\n \r\n       Radiologist: " + getrad + "\r\n \r\n       Radiologist Reason For Image Flag: "
                             + getreason + "\r\n \r\n       Radiologist Notes (not always included by radiologist): "
                             + getnotes + "\r\n \r\nPlease, contact your supervisor if you have any further questions."
                             "  \r\n \r\nIf you received this message by mistake, contact physics@sanfordhealth.org."
                             " \r\n \r\nThank you, \r\n \r\nImaging Physics \r\nphysics@sanfordhealth.org")

# Move file after all functions have executed
files.copy(py.path.local(r"W:\SHARE8 Physics\Software\python\scripts\clahn\EPIC Rad Feedback\archived_daily_reports_unfiltered"))
os.remove(files)
os.remove(dailyreportname)
