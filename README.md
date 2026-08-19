# Automated Job Application & Follow-Up Tracker

## 📌 Project Overview

Automated Job Application & Follow-Up Tracker is a Google Apps Script and Google Sheets based automation system designed to simplify the job application process.

The system helps manage job application records, send personalized application emails, attach PDF resumes, track recruiter responses, prevent duplicate applications, and send scheduled follow-up emails.

The project demonstrates practical workflow automation using Google Apps Script, Gmail, Google Sheets, and Google Drive.

---

## ✨ Features

- Automated job application emails
- Personalized email content
- PDF resume attachment
- Daily email sending limit
- Duplicate email protection
- Application status tracking
- Sent date tracking
- Automatic reply detection
- Reply date and reply details tracking
- Weekly follow-up automation
- Follow-up count tracking
- Last follow-up date tracking
- Google Sheets based application management

---

## 🛠️ Technologies Used

- Google Apps Script
- JavaScript
- Google Sheets
- Gmail
- Google Drive
- CSV
- Google Apps Script Triggers

---

## ⚙️ How It Works

### 1. Application Data

Recruiter and company information is maintained in a structured Google Sheet.

The system reads:

- Name
- Email
- Job Title
- Company
- Application Status

### 2. Automated Email

The `Code.gs` script reads the application records and sends personalized emails.

A PDF resume can be attached automatically from Google Drive.

### 3. Duplicate Protection

Before sending an email, the system checks whether the same email address already exists earlier in the spreadsheet.

Duplicate records are skipped to help prevent repeated applications.

### 4. Reply Tracking

The `ReplyTracker.gs` script searches Gmail for responses from tracked email addresses.

When a reply is detected, the spreadsheet can be updated with:

- Reply status
- Reply date
- Reply details
- Result

### 5. Weekly Follow-Up

The `FollowUp.gs` script handles follow-up communication.

Applications that have not received a reply can be considered for follow-up after the configured waiting period.

The system also records:

- Follow-up count
- Last follow-up date

---

## 📊 Spreadsheet Structure

| Column | Description |
|---|---|
| Name | Recruiter or contact name |
| Email | Contact email |
| Title | Job/recruitment title |
| Company | Company name |
| Sent Date | Date application was sent |
| Reply | Reply status |
| Result | Application result |
| Reply Date | Date of received reply |
| Reply Details | Short reply content |
| Follow-up Count | Number of follow-ups |
| Last Follow-up | Most recent follow-up date |

---

## 📁 Project Structure

```text
Automated-Job-Application-Follow-Up-Tracker/
│
├── Code.gs
├── ReplyTracker.gs
├── FollowUp.gs
├── sample-data.csv
├── Sample_Resume_Automated_Job_Application_Project.pdf
└── README.md
