import os
import pandas as pd
import pyodbc
import win32com.client as win32
from datetime import datetime
import py

# Set the folder path to search for Excel files
folder_path = 'path to file'

# Set the output file path for the 'daily' table archive
today_date_str = datetime.today().strftime('%Y-%m-%d-%H-%M-%S')
output_path = 'path to file'

# Set the output file path for the 'fargo' table archives
output_path1 = 'path to file'
output_path2 = 'path to file'


# Set the email addresses of the recipients for the daily report
recipients = []

# Used for testing to only send daily report to me
# recipients = ['christopher.lahn@sanfordhealth.org', ]

# Create dictionary to map technologist name to email address
email_dict = {}

# open the file(s) to process in the incoming folder and resave so it doesn't have a password on it.
dirname = py.path.local('path to file')
for f in dirname.visit(fil='*.xlsx', bf=True):
    try:
        # This will unprotect workbook and save it again.
        xcl = win32.Dispatch('Excel.Application')
        pw_str = ''
        wb = xcl.Workbooks.Open(f, False, False, None, pw_str)
        wb.Password = ""
        xcl.DisplayAlerts = False
        wb.Save()
        xcl.Quit()
    except:
        pass

# Connect to Microsoft Access database
conn = pyodbc.connect('path to file')

# Create cursor to execute SQL commands
cursor = conn.cursor()

# Clear the 'daily' table
cursor.execute('DELETE FROM daily')

# Loop through each file in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        # Read Excel file into a pandas DataFrame
        file_path = os.path.join(folder_path, file_name)
        # this keeps leading zero so the duplicates match up in database.
        df = pd.read_excel(file_path, dtype={'Accession #': str})
        # set column headers to match database headers
        df.columns = [
                    'mrn', 'accession', 'begin_exam', 'quality_user',
                    'quality_element', 'quality_comment', 'technologist',
                    'dept', 'category', 'procedure'
                    ]
        df['accession'] = df['accession'].astype('object')
        df = df.astype(str)
        # print(df.dtypes)
        # Loop through each row in the DataFrame
        for i, row in df.iterrows():
            # Check if 'accession' value is unique in the 'fargo' table
            cursor.execute("SELECT COUNT(*) FROM fargo WHERE accession = ?", (row['accession'],))
            count = cursor.fetchone()[0]
            if count == 0:
                # Append data to 'fargo' table
                cursor.execute("INSERT INTO fargo VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (row['mrn'], row['accession'], row['begin_exam'], row['quality_user'], row['quality_element'], row['quality_comment'], row['technologist'], row['dept'], row['category'], row['procedure']))
                conn.commit()

                # Append data to 'daily' table
                cursor.execute("INSERT INTO daily VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (row['mrn'], row['accession'], row['begin_exam'], row['quality_user'], row['quality_element'], row['quality_comment'], row['technologist'], row['dept'], row['category'], row['procedure']))
                conn.commit()


 
# Export 'daily' table to Excel
daily_df = pd.read_sql_query("SELECT * FROM daily", conn)
daily_df.to_excel(output_path, index=False)

# Export 'fargo' table to Excel in two locations
fargo_df = pd.read_sql_query("SELECT * FROM fargo", conn)
fargo_df.to_excel(output_path1, index=False)
# fargo_df.to_excel(output_path2, index=False)

#Copy master to SharePoint
fileMaster = py.path.local('path to file')

try:
    fileMaster.copy(py.path.local('path to file'))
    fileMaster.copy(py.path.local('path to file'))

except Exception as err:
    print("you may have not been signed in.")
    pass

# Create a copy of the DataFrame
daily_df_copy = daily_df.copy()

# Drop the 'mrn' column from the DataFrame copy
daily_df_copy = daily_df_copy.drop(columns=['mrn'])

def style_html_table(html_table):
    table_start = html_table.find('<table')
    table_end = html_table.find('>') + 1
    styled_table = html_table[:table_start] + '<table style="border-collapse: collapse; font-family: Arial, sans-serif;">' + html_table[table_end:]

    styled_table = styled_table.replace('<th', '<th style="border: 1px solid black; padding: 5px; font-weight: bold;"')
    styled_table = styled_table.replace('<td', '<td style="border: 1px solid black; padding: 5px;"')

    return styled_table

html_table = daily_df_copy.to_html(index=False, border=0)
html_table = style_html_table(html_table)

# If no feedback (empty daily_df) send no feedback message. Otherwise send daily report.
if daily_df.empty:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ';'.join(recipients)
    mail.Subject = 'AUTOMATED MESSAGE: No Radiologist Feedback Today'
    mail.Body = 'There was no radiologist feedback left today.'
    mail.Send()
else:
    # Email 'daily' table as an HTML table in the email body to each recipient using Outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ';'.join(recipients)
    mail.Subject = 'AUTOMATED MESSAGE: Radiologist Feedback Daily Report'
    mail.HTMLBody = f'''Please find the daily report below:
                        <br><br>
                        {html_table}
                        <br><br>
                        You can find the master file with all feedback here: https://tinyurl.com/bd9phx3j
                        <br><br>
                        Contact imaging physics if you have any questions. physics@sanfordhealth.org
                        '''
    mail.Send()


# Email each row of 'daily' table to the technologist listed in the row

# Map for renaming columns in df_daily to read easier in email body
col_name_map = {
    'begin_exam': 'Exam Date',
    'accession': 'Accession Number',
    'quality_element': 'Reason',
    'quality_comment': "Radiologist's Comment",
    'technologist': 'Technologist',
    'dept': 'Department',
    'category': 'Category',
    'procedure': 'Procedure',
    'quality_user': 'Radiologist',
}

for i, row in daily_df.iterrows():
    technologist_name = row['technologist']
    if pd.isna(technologist_name):
        # Set technologist email to christopher.lahn@sanfordhealth.org
        technologist_email = "christopher.lahn@sanfordhealth.org"
    else:
        # Check if Technologist is in the email_dict
        if technologist_name in email_dict:
            # Send email to the address in the dictionary
            technologist_email = email_dict[technologist_name]
        else:
            # Construct the email recipient from the technologist name
            name_parts = technologist_name.split(',')
            last_name = name_parts[0].strip()
            first_name = name_parts[1].strip().split()[0]
            technologist_email = f"{first_name}.{last_name}@sanfordhealth.org"

        # Check if Quality Element includes "Physics"
    if "Physics" in row['quality_element']:
        # Send separate email to physics@sanfordhealth.org
        mail = outlook.CreateItem(0)
        mail.To = ""
        mail.Subject = 'Automated Message:  Image Quality Feedback'
        mail.Body = ''
        
        # Add the message to the email body
        mail.Body += ("Hello,\n\nThis is an automated message. No reply is necessary. "
                    "\n\nAn image that you completed in Radiant was flagged for image quality review by a radiologist."
                    "\nPlease, use the accession number provided here to look up your exam in PACS and review the image along with the radiologist's feedback provided in this email."
                    "\n\nNotes:\n1. Please, respond via email to the radiologist listed below if they requested so in the Radiologist Notes section."
                    "\n2. nan = no value entered by Radiologist."
                    "\n3. If Radiologist Reason for Image Flag includes or is only Physics, the exam was flagged for possible equipment issues and will be reviewed by Imaging Physics, as well. Contact Imaging Physics if you have any input which may help us resolve the potential equipment issue."
                    )

        # Add each column of data to the email body
        for col_name, col_val in row.iteritems():
            mail.Body += f"{col_name}: {col_val}\n"

        # Add the message to the email body
        mail.Body += ("\nPlease, contact your supervisor if you have any further questions. \nThank you, "
        "\nIf you received this message by mistake, contact physics@sanfordhealth.org. \nImaging Physics \nphysics@sanfordhealth.org")
        
        mail.Send()
    else:
        # Create email with row data in the body
        mail = outlook.CreateItem(0)
        mail.To = technologist_email
        mail.Subject = 'Automated Message:  Image Quality Feedback'
        mail.Body = ''
        # Add the message to the email body
        mail.Body += ("Hello,\n\nThis is an automated message. No reply is necessary. "
                    "\n\nAn image that you completed in Radiant was flagged for image quality review by a radiologist."
                    "\nPlease, use the accession number provided here to look up your exam in PACS and review the image along with the radiologist's feedback provided in this email."
                    "\n\nNotes:\n1. Please, respond via email to the radiologist listed below if they requested so in the Radiologist Notes section."
                    "\n2. nan = no value entered by Radiologist."
                    "\n3. If Radiologist Reason for Image Flag includes or is only Physics, the exam was flagged for possible equipment issues and will be reviewed by Imaging Physics, as well. Contact Imaging Physics if you have any input which may help us resolve the potential equipment issue."
                    )

        # Add each column of data to the email body
        for col_name, col_val in row.iteritems():
            # do not add mrn column to email body
            if col_name != 'mrn':
                # Check if the column name needs to be mapped to a new one
                if col_name in col_name_map:
                    col_name = col_name_map[col_name]
            
                mail.Body += f"{col_name}: {col_val}\n"

        # Add footer to the message to the email body
        mail.Body += ("\nPlease, contact your supervisor if you have any further questions. \nThank you, "
        "\nIf you received this message by mistake, contact physics@sanfordhealth.org. \nImaging Physics \nphysics@sanfordhealth.org")

        
        mail.Send()

# Delete the Excel files in the incoming folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)
        os.remove(file_path)

# Commit changes and close connection
conn.commit()
cursor.close()
conn.close()

print('Emails sent successfully')
