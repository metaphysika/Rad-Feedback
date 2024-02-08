# Radiologist Feedback Automation

## Overview

This script automates the process of distributing radiologist feedback from reports downloaded from Epic. It emails the comments to the respective staff members and sends a daily report to supervisors. Additionally, it archives all comments in a database for future analysis.

## Features

    - Processes incoming daily reports containing radiologist feedback.
    - Saves a copy of the processed data for archival purposes.
    - Emails feedback directly to staff based on flags in the report.
    - Sends a compiled daily feedback report to supervisors.
    - Archives all feedback in a Microsoft Access database for analysis.

## Prerequisites

    - Python 3.x
    - Microsoft Access
    - Windows environment with Outlook for sending emails.
    - pandas, pyodbc, win32com.client, datetime, and py Python libraries.

## Installation

Ensure you have the necessary Python libraries installed. You can install these using pip:

pip install pandas pyodbc pypiwin32

py is a part of pylib and can also be installed via pip if not already present:

pip install py

## Configuration

Before running the script, configure the following paths and variables:


    - `folder_path`: Directory containing the incoming Excel reports.
    - `output_path`, `output_path1`, `output_path2`: Paths for saving the daily and master table archives.
    - `recipients`: List of supervisor email addresses for the daily report.
    - `email_dict`: Dictionary mapping technologist names to their email addresses.


## Running the Script

Execute the script in a Python environment. Ensure that you have read and write permissions for the specified directories and the necessary access to the email server for sending emails.
How It Works


    - The script searches for Excel files in the specified folder and processes them.
    - It unprotects and saves the files if they are password-protected.
    - It connects to an Access database and updates tables with the new data.
    - The script then sends personalized emails to staff with their feedback.
    - A compiled report is emailed to supervisors.
    - The script also handles copying the master file to a SharePoint directory.
    - It styles the daily feedback data as an HTML table for email body inclusion.


## Output

    - The script will display a confirmation in the console once the emails have been sent successfully.
    - Excel files from the incoming folder are deleted after processing.
    - The Access database is updated with the new data.


## Notes

    - The script must be run in an environment where Outlook is installed and configured.
    - Ensure that the email sending functionality adheres to your organization's IT policies.
    - Regular maintenance and verification of the database connections and file paths are recommended.

Disclaimer

This script is provided as-is, and the author is not responsible for any unintended data loss or email miscommunication. Use it at your own risk, and ensure proper testing is done in a controlled environment before deploying it in a production setting.