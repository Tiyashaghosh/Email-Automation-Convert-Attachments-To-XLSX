# 📩 Email-Automation-Convert-Attachments-To-XLSX


**Automate your email workflow** — This Python script connects to your Gmail inbox via IMAP, searches for emails with attachments (CSV, HTML tables, or XLS), converts them into `.xlsx` format, and sends a reply email with the converted file.

---

## 🚀 Features

- ✅ Connects securely to Gmail using IMAP
- ✅ Downloads email attachments to a local folder
- ✅ Detects file type using `python-magic`
- ✅ Converts:
  - `.csv` → `.xlsx`
  - `.html` (tables) → `.xlsx`
  - `.xls` → `.xlsx`
- ✅ Automatically replies to the sender with the converted `.xlsx` file
- ✅ Uses environment variables for secure credentials
