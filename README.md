# NYS_APPLY
# NY State Jobs Automator v3

An AI-powered automation tool that scrapes NY State job listings, analyzes job postings, generates tailored cover letters, and automatically prepares application packages — including Outlook draft emails for review before sending.

---

## 🚀 Features

### 🔍 Job Scraping
- Uses **Selenium (real browser automation)** to scrape NY State job listings
- Supports pagination and full vacancy extraction
- Opens each job posting and extracts all tabbed content:
  - Basics
  - Schedule
  - Location
  - Job Specifics
  - How to Apply

---

### 🧠 AI-Powered Processing (Ollama / LLaMA3)
- Automatically determines:
  - Email apply / portal apply / fax / mail-only
- Selects the **best matching resume**
- Generates a **tailored professional cover letter**
- Extracts structured application instructions

---

### 📄 Document Generation
Each job gets its own folder containing:
- 📄 Tailored Cover Letter (PDF, Times New Roman, formatted letterhead)
- 📄 Selected Resume
- 📄 Transcript (if required by job posting)
- 📄 HOW_TO_APPLY.txt (structured breakdown)
- 📄 job_summary.json (metadata + decisions)

---

### ✉️ Outlook Integration (IMPORTANT)
- Automatically creates **Outlook Draft Emails (NOT SENT)**
- Uses Windows Outlook desktop via `win32com`
- Attaches:
  - Resume
  - Cover letter PDF
  - Transcript (if required)
- You manually review and send the email

✔ Safe by design — nothing is auto-sent

---

### 📊 Final Reports
At the end of execution:
- `EMAIL_APPLY_JOBS.txt` → jobs with email drafts
- `OTHER_APPLY_JOBS.txt` → portal/fax/mail jobs
- Organized per-job folders for full traceability

---

## ⚙️ Requirements

### Python Dependencies
Install everything using:

```bash
pip install -r requirements.txt
