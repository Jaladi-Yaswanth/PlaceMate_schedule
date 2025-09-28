#  Automated Placement Mail → Google Calendar Integration

## 📌 Problem Statement
University placement teams often send **emails with attachments (Excel/Sheets)** listing candidates, venues, and schedules for tests or talks.  
Manually, students must:  
- Open emails  
- Download attachments  
- Search for their name/registration number  
- Note the venue/date  
- Create calendar events manually  

⚡ This is time-consuming.

**Solution:** An automated Gmail + Google Apps Script system that:  
- Reads unread mails from placement office mailing lists  
- Extracts details (date, time, venue) from mail subject/attachments  
- Checks if the logged-in student is listed  
- Automatically creates a **Google Calendar Event** with correct details 🎯

---

## 🛠️ Tech Stack
- **Google Apps Script** – automation engine   
- **Drive API** – handle Excel/Sheets conversion  
- **SpreadsheetApp API** – parse attachments  
- **CalendarApp API** – auto-create calendar events  

---


## ⚙️ Setup Instructions

Follow these steps to set up and run the Placement Automation script:

1. **Create a Google Apps Script Project**  
   - Go to [Google Apps Script](https://script.google.com/).  
   - Click **New Project**.  
   - Give your project a meaningful name, e.g., `Placement Event Automation`.

2. **Add Your Script Code**  
   - Open the `Code.gs` file in the project editor.  
   - Copy and paste the full script code provided in this repository.  
   - Make sure to include all functions (`myFunction`, `extractDateTimeFromText`, `extractvenue`, etc.).

3. **Enable Advanced Google Services**  
   - Go to `Services` in the left sidebar → click **+ Add a service**.  
   - Select **Drive API** → Click **Add**.  
   - This allows the script to convert and read Excel files (`.xlsx`) from Gmail attachments.

4. **Replace Personal Placeholders**  
   - Update the following variables in your script:  
     ```javascript
     const my_name = "Your_Name";          // Replace with your full name
     const my_reg_number = "Your_Registration_Number";  // Replace with your registration number
     ```
   - These are used to check if your name appears in the attached sheets.

5. **Set Up a Trigger to Run Automatically**  
   - Click the **Triggers (clock icon)** in Apps Script → **Add Trigger**.  
   - Select the function `myFunction` to run.  
    ### Trigger Configuration
   - **Event Source** → `Time-driven`  
   - **Type of Trigger** → `Minutes timer`  
    - **Interval** → `Every 30 minutes` (since it will check in last 30 min window).
   - This ensures the script regularly checks your Gmail for new placement or test emails.  

6. **Authorize the Script**  
   - When you run the script or set up the trigger, Google will ask for permissions.  
   - Grant the following access:  
     - **Gmail** → Read emails and attachments.  
     - **Drive** → Create and trash temporary files for Excel/Sheet conversion.  
     - **Calendar** → Create events automatically.  
   - This allows the script to function fully without manual intervention.  

7. **Test the Script**  
   - Send a sample email to your Gmail with your name in the attached Excel file.  
   - Make sure the subject contains keywords like `talk` or `test`.  
   - Run the script manually once to confirm it correctly creates a calendar event.  
   - Check the logs in Apps Script (`View → Logs`) for debug information.

