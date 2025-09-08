# 📧 Automated Placement Mail → Google Calendar Integration

## 📌 Problem Statement
University placement teams often send **emails with attachments (Excel/Sheets)** listing candidates, venues, and schedules for tests or talks.  
Manually, students must:  
- Open emails  
- Download attachments  
- Search for their name/registration number  
- Note the venue/date  
- Create calendar events manually  

⚡ This is time-consuming, error-prone, and inefficient.

**Solution:** An automated Gmail + Google Apps Script system that:  
- Reads unread mails from placement office mailing lists  
- Extracts details (date, time, venue) from mail subject/attachments  
- Checks if the logged-in student is listed  
- Automatically creates a **Google Calendar Event** with correct details 🎯

---

## 🛠️ Tech Stack
- **Google Apps Script** – automation engine  
- **GmailApp API** – read and filter emails  
- **Drive API** – handle Excel/Sheets conversion  
- **SpreadsheetApp API** – parse attachments  
- **CalendarApp API** – auto-create calendar events  

---

## ⚙️ Setup Instructions
1. Open [Google Apps Script](https://script.google.com/)  
2. Create a new project and paste the code from `Code.gs`  
3. Enable **Google Drive API** in:  
   - Apps Script → Services → Add Drive API  
4. Replace placeholders:  
   ```javascript
   const my_name = "Your_Name";
   const my_reg_number = "Your_Registration_Number";
