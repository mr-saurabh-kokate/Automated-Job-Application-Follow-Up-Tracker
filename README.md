# Automated Job Application & Follow-Up Tracker

## Project Overview

Automated Job Application & Follow-Up Tracker is a Google Apps Script based automation system designed to simplify and manage the job application process.

The project uses Google Sheets as a central application tracker and integrates Gmail and Google Drive to automate personalized job application emails, PDF resume attachments, reply tracking, and follow-up communication.

The main goal is to reduce repetitive manual work involved in sending job applications and maintaining recruiter communication records.

## How It Works

Recruiter and company information is maintained in a structured Google Sheet. The automation reads the required information and sends personalized application emails with a PDF resume attachment.

After an email is successfully sent, the system records the sent date and application status in the spreadsheet.

The system also includes duplicate email protection to prevent repeated applications being sent to the same email address.

## Reply Tracking

The system periodically checks Gmail for responses from recruiters or companies.

When a reply is detected, the spreadsheet is automatically updated with:

- Reply status
- Reply date
- Reply details
- Result status

This makes it easier to identify companies that have responded without manually checking every application.

## Follow-Up Automation

For applications where no reply has been received, the system supports scheduled follow-up emails.

Weekly follow-ups can be automated while maintaining follow-up limits. Once a recruiter responds, further follow-ups can be avoided.

## Key Features

- Automated personalized job application emails
- PDF resume attachment
- Daily email limit
- Duplicate email protection
- Sent date tracking
- Application status tracking
- Automated reply detection
- Reply date tracking
- Reply details tracking
- Weekly follow-up automation
- Google Sheets based tracking
- Scheduled triggers

## Technologies Used

- Google Apps Script
- JavaScript
- Gmail
- Google Sheets
- Google Drive

## Project Structure

```text
├── Code.gs
├── ReplyTracker.gs
├── FollowUp.gs
├── sample-data.csv
├── sample-resume.pdf
└── README.md
