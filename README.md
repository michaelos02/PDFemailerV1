# Personalized Email Sender for Google Sheets

This Google Apps Script project transforms your Google Sheet into a powerful personalized email sending tool, complete with dynamic field merging, attachment capabilities, and robust status tracking. It's designed to make sending personalized emails to a list of recipients as easy and reliable as possible.

---

## Features

* **Custom Google Sheet Integration:** A custom menu "Email Tools" is added directly to your Google Sheet for easy access to all functionalities.
* **Intuitive Email Composer:** A dedicated dialog allows you to craft and save email subject lines and HTML bodies with rich text editing capabilities.
* **Dynamic Field Merging:** Seamlessly insert data from your Google Sheet columns (e.g., `[[First Name]]`, `[[Email]]`) directly into your email subject and body.
* **Google Doc Attachments:** Attach Google Docs from your Drive to your emails, which are automatically converted to PDF format upon sending.
* **Robust Email Status Tracking:** A dedicated "Email Status" column is automatically managed in your sheet, providing real-time updates on each email's delivery.
    * **Red Error Highlighting:** Failed sends or attachment issues are clearly marked in **red text** within the status column for easy identification.
* **Mandatory Email Preview:** Before sending, a preview dialog displays the personalized email for the first row of your data, ensuring accuracy and preventing errors.
* **Smart Resend/Skip Logic:** The sender intelligently skips rows that already have a status in the "Email Status" column. This allows you to easily resend emails by simply clearing the status from specific cells, mimicking the behavior of tools like FormMule.

---

## How It Works

1.  **Prepare Your Google Sheet:**
    * Ensure your active sheet has a header row.
    * Include an **`Email`** column for recipient addresses.
    * (Optional) Include a **`DocLink`** column with Google Doc IDs or shareable URLs if you wish to attach documents.
    * Populate your rows with data you want to use for personalization (e.g., `First Name`, `Company`, `Amount`).

2.  **Access the Email Composer:**
    * Open your Google Sheet.
    * Go to `Email Tools > Open Email Composer` in the top menu.
    * Write your email subject and body. Use `[[Column Name]]` to insert dynamic fields.
    * Click "Save Template". This stores your template in a hidden sheet named `EmailComposerTemplate`.

3.  **Send Personalized Emails:**
    * Go to `Email Tools > Send Personalized Emails` in the top menu.
    * A **preview dialog** will appear, showing you exactly how the email for your *first data row* will look.
    * Review the preview carefully.
    * Click **"Send All Emails"** to proceed with sending to all eligible recipients in your sheet.
    * Click **"Cancel Sending"** if you need to make adjustments.

4.  **Monitor Status:**
    * An "Email Status" column will be automatically created (if it doesn't exist) in your active sheet.
    * This column will update with the sending status for each row (e.g., "Email sent on...", "Error: Invalid email", "Skipped: Status already present").

---

## Getting Started (Installation)

1.  **Open your Google Sheet:** Go to [sheets.new](https://sheets.new) or open an existing Google Sheet where you want to use this tool.
2.  **Open the Apps Script Editor:** Click `Extensions > Apps Script` in the Google Sheet menu. This will open a new browser tab with the Apps Script editor.
3.  **Create `Code.gs`:**
    * In the Apps Script editor, if there's a default `Code.gs` file, select all its content and delete it.
    * Paste the entire content of the provided `Code.gs` into this file.
    * Save the file (`Ctrl + S` or `Cmd + S`).
4.  **Create `index.html`:**
    * In the Apps Script editor, click `File > New > HTML file`.
    * Name the new file `index.html`.
    * Paste the entire content of the provided `index.html` into this file.
    * Save the file.
5.  **Create `preview.html`:**
    * In the Apps Script editor, click `File > New > HTML file`.
    * Name the new file `preview.html`.
    * Paste the entire content of the provided `preview.html` into this file.
    * Save the file.
6.  **Refresh Your Google Sheet:** Go back to your Google Sheet tab and perform a **hard refresh** (`Ctrl + F5` on Windows/Linux, `Cmd + Shift + R` on Mac). You should now see an "Email Tools" menu item in your sheet's menu bar.
7.  **Authorize the Script:** The first time you use any custom menu item, Google will prompt you to authorize the script. Follow the on-screen instructions:
    * Click "Review permissions".
    * Select your Google account.
    * Click "Allow" (you may need to click "Advanced" and "Go to [Project Name] (unsafe)" if you haven't published the script). The script needs permission to send emails on your behalf and access your Google Sheets and Drive.

---

## Code Structure

* `Code.gs`: Contains the core Apps Script logic, including menu creation, template handling, data processing, and email sending functions.
* `index.html`: Provides the user interface for the Email Composer dialog.
* `preview.html`: Provides the user interface for the email preview dialog before sending.


