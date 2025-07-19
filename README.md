# Bulk Email Sender with Resume Attachment

This is a **test automation script** designed to send emails in bulk using an Excel sheet that contains recipient email addresses in the first column. The script supports sending PDF resume attachments and uses Gmail SMTP for email delivery.

> ⚠️ **Note**: This script is for learning and testing purposes. Avoid spamming or violating email service terms.

---
---

## 🚀 What This Script Does

- Reads email addresses from the first column of an Excel file.
- Filters out blacklisted domains (e.g., `deqode.com`, `cis.com`).
- Sends customized emails with a PDF attachment using Gmail SMTP.
- Uses multithreading (`ThreadPoolExecutor`) for faster email dispatch.
- Creates `failed_emails.xlsx` containing emails that failed to send.

---

## 🔧 Setup Instructions

### 1. Clone the Repository

```bash
git clone https://github.com/Devendra176/scripts.git
cd scripts
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```
### 3. Create .env File
```bash
# Production
SENDER=your@gmail.com
PASSWORD=YourAppPassword
HOST=smtp.gmail.com
PORT=587

# Test (e.g. Mailtrap)
TEST_SENDER=yourId
TEST_PASSWORD=YourPassword
TEST_HOST=sandbox.smtp.mailtrap.io
TEST_PORT=587
```
### 4. Prepare Email Content
- In global_var.py, define: according to your requirement
```bash
SUBJECT = "Job Application for Python Developer Role"
BODY = "Dear Hiring Team,\n\nPlease find my resume attached.\n\nRegards,\nDevendra Lodhi"
EMAIL_FILE = "emails.xlsx"
ATTACHMENT = "resume.pdf"
```
### 5. Prepare Excel Sheet
- Create an Excel file named emails.xlsx with one email per row in the first column:
```bash
| Email             |
|-------------------|
| hr@company1.com   |
| careers@xyz.com   |
| example@gmail.com |
```
## ▶️ Running the Script
- Run in Production Mode:
```bash
python send_mail_recruters.py --prod --thread 3 --delay 2
```
- Run in Test Mode (e.g., with Mailtrap):
```bash
python send_mail_recruters.py --test --thread 3 --delay 2
```
## ✅ Output
- ✅ Emails sent will be shown in the console.

- ❌ Failed emails will be saved in failed_emails.xlsx.

## 🧵 Multithreading
- The script uses concurrent.futures.ThreadPoolExecutor to parallelize the sending process. You can configure the number of threads for speed.


## 🙋‍♂️ Contact
- For questions or contributions:
Devendra Lodhi
📧 [devendralodhi176@gmail.com]
📞 +91-9617128605


- Let me know if you want to include screenshots, GitHub badges, or make it more personalized!















