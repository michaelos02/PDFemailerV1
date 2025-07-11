## Google Apps Script: Personalized Email Composer

This project provides a custom Email Composer built with Google Apps Script, designed to streamline sending personalized emails directly from a Google Sheet. It features a rich text editor, dynamic field insertion based on spreadsheet headers, and the ability to attach Google Docs as PDFs.

### ✨ Key Features

* **Rich Text Editor:** Compose email content with standard formatting (bold, italic, lists, colors, etc.).
* **Dynamic Field Insertion:** Easily insert `[[Header Name]]` placeholders into your subject and body, which are automatically replaced with data from your Google Sheet rows during sending.
* **Google Sheet Integration:** Uses data from your active Google Sheet to personalize emails.
* **Google Drive Attachments:** Supports attaching Google Docs (converted to PDF) via links in a dedicated spreadsheet column.
* **Intuitive UI:** A custom dialog launched directly from your Google Sheet for editing email templates.

### 🚀 Setup & Installation

To use this Email Composer in your Google Workspace environment:

1.  **Create a new Google Sheet** or open an existing one that will contain your recipient data.
2.  Open the **Apps Script editor** from your Google Sheet: Go to `Extensions > Apps Script`.
3.  In the Apps Script editor, you will see a default `Code.gs` file.
4.  Create three new HTML files:
    * Click `File > New > HTML file` and name them `index.html`, `index_CSS.html`, and `index_JS.html` respectively.
5.  **Copy the code** from the corresponding files in this GitHub repository (`Code.gs`, `index.html`, `index_CSS.html`, `index_JS.html`) into the newly created Apps Script files.
6.  **Save all files** in the Apps Script editor.
7.  **Return to your Google Sheet** and refresh the page. A new "Email Tools" custom menu will appear.

### 📊 Spreadsheet Data Structure

For the email sending functionality to work correctly, ensure your active Google Sheet is structured with a **header row** (first row) containing meaningful column names. At a minimum, you **must have a column with the header `Email`** (case-insensitive).

Optionally, for document attachments, include a column with the header `DocLink` (case-insensitive) containing Google Doc URLs or IDs.

### 💡 How It Works

Upon opening your Google Sheet, an `Email Tools` custom menu is created.
* **`Open Email Composer`**: Launches a custom dialog where you can compose and save your email subject and body template. These templates are stored in a hidden sheet within your spreadsheet.
* **`Send Personalized Emails`**: Processes the data in your active sheet, replaces placeholders in your saved template with row-specific data, and sends individual emails.

---
