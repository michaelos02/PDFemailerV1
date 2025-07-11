/**
 * @file Code.gs
 * @description Google Apps Script functions for the Email Composer.
 * This script manages custom menu creation, dialog display,
 * template saving/loading, and personalized email sending with attachments.
 */

// --- Configuration Constants ---
const TEMPLATE_SHEET_NAME = 'EmailComposerTemplate'; // Name of the hidden sheet for storing email templates
const DOC_LINK_HEADER_NAME = 'DocLink'; // Standard header name for the Google Doc ID/URL column in the data sheet

// Cell locations within the TEMPLATE_SHEET_NAME for storing template parts
const SUBJECT_CELL = 'B1';
const BODY_CELL = 'B2';

// --- Spreadsheet UI Functions ---

/**
 * Creates a custom menu in the Google Sheet's UI when the spreadsheet is opened.
 * This function is automatically triggered by the 'onOpen' event.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Tools') // Changed menu name for clarity
      .addItem('Open Email Composer', 'showEmailComposerDialog') // Function to open the editor dialog
      .addSeparator() // Adds a visual separator in the menu
      .addItem('Send Personalized Emails', 'sendPersonalizedEmails') // Function to trigger email sending
      .addToUi();
}

/**
 * Displays the custom Email Composer in a modal dialog within the Google Sheet.
 * The HTML content for the dialog is loaded from 'index.html'.
 */
function showEmailComposerDialog() {
  const htmlOutput = HtmlService.createTemplateFromFile('index.html')
      .evaluate()
      .setWidth(850)
      .setHeight(600)
      .setTitle('Email Composer'); // Sets the title of the dialog window

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, htmlOutput.getTitle());
}

// --- Helper Functions ---

/**
 * Helper function to include external HTML, CSS, or JavaScript files
 * into a main HTML template file using Apps Script scriplets.
 *
 * @param {string} filename The name of the Apps Script HTML file to include (e.g., 'index_CSS').
 * @returns {string} The content of the specified file as an HTML string.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Retrieves a sheet by its name, creating it if it does not already exist.
 * If a new sheet is created, it is hidden and resized to a minimal size.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The active spreadsheet object.
 * @param {string} sheetName The name of the sheet to retrieve or create.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Sheet object.
 * @throws {Error} If the sheet cannot be accessed or created due to permissions or other issues.
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    try {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.hideSheet(); // Hide the newly created sheet from user view

      // Resize the new sheet to a minimal size (e.g., 10 rows, 4 columns)
      const maxRows = sheet.getMaxRows();
      const maxColumns = sheet.getMaxColumns();

      if (maxColumns > 4) {
        sheet.deleteColumns(5, maxColumns - 4);
      }
      if (maxRows > 10) {
        sheet.deleteRows(11, maxRows - 10);
      }

      Logger.log(`Sheet "${sheetName}" created, hidden, and resized to 10 rows and 4 columns.`);
    } catch (e) {
      throw new Error(`Could not create sheet "${sheetName}". Please check permissions or sheet name: ${e.message}`);
    }
  }
  return sheet;
}

/**
 * Validates if a given string is a basic email address format.
 * This is a simple regex and not exhaustive for all valid email addresses.
 *
 * @param {string} email The email string to validate.
 * @returns {boolean} True if the email string matches a basic email format, false otherwise.
 */
function validateEmail(email) {
  const emailRegex = /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return emailRegex.test(String(email).toLowerCase());
}

/**
 * Extracts a Google Doc ID from a full Google Drive URL or returns the ID
 * if the input string is already a direct Google Doc ID.
 *
 * @param {string} docLink The Google Doc URL or a direct Google Doc ID.
 * @returns {string|null} The extracted Google Doc ID, or null if the input is not a valid link or ID.
 */
function extractDocId(docLink) {
  if (!docLink) return null;

  // Regex to match the document ID from a standard Google Docs URL
  const urlRegex = /document\/d\/([a-zA-Z0-9_-]+)/;
  const match = docLink.match(urlRegex);

  if (match && match[1]) {
    return match[1]; // Return the ID found in the URL
  } else if (docLink.length > 20 && !docLink.includes('/')) {
    // Simple heuristic: if it's long and doesn't contain slashes, assume it's a direct ID
    return docLink;
  }
  return null; // Not a recognized ID or URL format
}

// --- Template Management Functions ---

/**
 * Saves the provided subject and HTML body content to the designated template sheet.
 * This function is called from the client-side JavaScript via `google.script.run`.
 *
 * @param {string} subject The subject line template string to save.
 * @param {string} body The HTML string content (email body) to save.
 */
function saveTemplate(subject, body) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(spreadsheet, TEMPLATE_SHEET_NAME);

    sheet.getRange(SUBJECT_CELL).setValue(subject);
    sheet.getRange(BODY_CELL).setValue(body);

    Logger.log(`Template (Subject & Body) saved successfully to "${TEMPLATE_SHEET_NAME}"!${SUBJECT_CELL} and ${BODY_CELL}.`);
  } catch (e) {
    Logger.log(`Error saving template: ${e.message}`);
    throw new Error(`Failed to save template: ${e.message}`);
  }
}

/**
 * Loads the subject and body content from the designated template sheet.
 * This function is called from the client-side JavaScript via `google.script.run`.
 *
 * @returns {Object} An object containing `subject` and `body` properties as strings.
 */
function loadTemplate() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(spreadsheet, TEMPLATE_SHEET_NAME);

    const subject = sheet.getRange(SUBJECT_CELL).getValue();
    const body = sheet.getRange(BODY_CELL).getValue();

    Logger.log(`Template (Subject & Body) loaded successfully from "${TEMPLATE_SHEET_NAME}"!${SUBJECT_CELL} and ${BODY_CELL}.`);
    return { subject: subject.toString(), body: body.toString() }; // Ensure values are returned as strings
  } catch (e) {
    Logger.log(`Error loading template: ${e.message}`);
    throw new Error(`Failed to load template: ${e.message}`);
  }
}

/**
 * Retrieves the header row (first row) from the currently active spreadsheet sheet.
 * This is used to populate the "Insert Field" dropdown in the client-side editor.
 *
 * @returns {string[]} An array of header names as they appear in the sheet.
 * @throws {Error} If no active sheet is found or columns are missing.
 */
function getSpreadsheetHeaders() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet(); // Get the sheet where the user's data resides
    if (!sheet) {
      throw new Error("No active sheet found in the spreadsheet.");
    }

    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      return []; // Return empty array if no columns exist
    }

    // Get all values from the first row, respecting cell formatting (e.g., "$100.00")
    const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
    Logger.log(`Headers retrieved: ${JSON.stringify(headers)}`);
    return headers;
  } catch (e) {
    Logger.log(`Error getting spreadsheet headers: ${e.message}`);
    throw new Error(`Failed to get spreadsheet headers: ${e.message}`);
  }
}

// --- Email Sending Logic ---

/**
 * Sends personalized emails based on the template stored in the hidden sheet
 * and data from the active spreadsheet sheet. Includes Google Doc attachments
 * converted to PDF.
 * This function is triggered via a custom menu item in the Google Sheet.
 */
function sendPersonalizedEmails() {
  const ui = SpreadsheetApp.getUi(); // UI object for displaying alerts to the user

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = spreadsheet.getActiveSheet(); // The sheet containing the user's data

    // 1. Retrieve Email Template (Subject and Body)
    const templateSheet = getOrCreateSheet(spreadsheet, TEMPLATE_SHEET_NAME);
    const emailSubjectTemplate = templateSheet.getRange(SUBJECT_CELL).getValue();
    const emailBodyTemplate = templateSheet.getRange(BODY_CELL).getValue();

    if (!emailSubjectTemplate && !emailBodyTemplate) {
      ui.alert('Email Template Missing', `Both the subject and body templates are empty in the "${TEMPLATE_SHEET_NAME}" sheet. Please create them using the Email Composer.`, ui.ButtonSet.OK);
      return;
    }

    // 2. Get All Data from the Active Sheet
    // Using getDisplayValues() to preserve formatting (e.g., currency symbols, dates)
    const allData = dataSheet.getDataRange().getDisplayValues();

    if (allData.length < 2) { // At least one header row and one data row are needed
      ui.alert('No Data Found', 'The active sheet must contain at least a header row and one row of data to send emails.', ui.ButtonSet.OK);
      return;
    }

    const headers = allData[0]; // First row contains the headers (field names)
    const dataRows = allData.slice(1); // All subsequent rows are the data for personalization

    // 3. Find Required Column Indices (case-insensitive search for headers)
    const emailColumnIndex = headers.findIndex(header => header.toLowerCase().trim() === 'email');
    const docLinkColumnIndex = headers.findIndex(header => header.toLowerCase().trim() === DOC_LINK_HEADER_NAME.toLowerCase());

    if (emailColumnIndex === -1) {
      ui.alert('Missing Email Column', 'The active sheet must have a column with the header "Email" to send emails.', ui.ButtonSet.OK);
      return;
    }
    // Note: The DocLink column is optional; emails will still send if it's missing or empty.

    let sentCount = 0;
    let failedCount = 0;
    const errors = []; // Collects detailed error messages for logging

    // 4. Iterate Through Each Data Row to Personalize and Send Email
    dataRows.forEach((row, rowIndex) => {
      // Calculate the actual row number in the spreadsheet (1-indexed, accounting for header)
      const currentRowNumber = rowIndex + 2;
      try {
        const recipientEmail = row[emailColumnIndex];

        // Basic validation for recipient email address
        if (!recipientEmail || !validateEmail(recipientEmail)) {
          const errorMessage = `Invalid or missing email address in row ${currentRowNumber}: '${recipientEmail}'. Skipping this row.`;
          Logger.log(errorMessage);
          errors.push(errorMessage);
          failedCount++;
          return; // Skip to the next row if email is invalid
        }

        // Initialize personalized subject and body with the templates
        let personalizedSubject = emailSubjectTemplate;
        let personalizedBody = emailBodyTemplate;

        // Replace all field codes (e.g., '[[Name]]') with data from the current row
        headers.forEach((header, colIndex) => {
          // Get field value; convert to string to handle various data types
          const fieldValue = row[colIndex] !== undefined && row[colIndex] !== null ? String(row[colIndex]) : '';
          // Construct the field code using the [[...]] format
          const fieldCode = `[[${header}]]`;
          // Create a global regular expression to replace all occurrences of the field code
          const regex = new RegExp(fieldCode.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'); // Escape special regex characters

          // Perform replacement in subject and body templates
          if (personalizedSubject) {
              personalizedSubject = personalizedSubject.replace(regex, fieldValue);
          }
          if (personalizedBody) {
              personalizedBody = personalizedBody.replace(regex, fieldValue);
          }
        });

        // 5. Handle Google Doc Attachment (if DocLink column exists and has a value)
        let attachments = [];
        if (docLinkColumnIndex !== -1) {
          const docLinkValue = row[docLinkColumnIndex];
          if (docLinkValue) { // Only proceed if there's a link/ID in the cell
            const docId = extractDocId(docLinkValue);
            if (docId) {
              try {
                const file = DriveApp.getFileById(docId);
                // Convert Google Docs to PDF for universal email client compatibility
                const pdfBlob = file.getAs(MimeType.PDF);
                attachments.push(pdfBlob);
                Logger.log(`Attached "${file.getName()}" (as PDF) for row ${currentRowNumber}.`);
              } catch (fileError) {
                const errorMessage = `Could not attach document for row ${currentRowNumber} (Link: '${docLinkValue}'): ${fileError.message}. Skipping attachment for this email.`;
                Logger.log(errorMessage);
                errors.push(errorMessage); // Log attachment-specific errors, but still try to send the email
              }
            } else {
              const errorMessage = `Invalid or unparseable document link in row ${currentRowNumber}: '${docLinkValue}'. Skipping attachment for this email.`;
              Logger.log(errorMessage);
              errors.push(errorMessage);
            }
          }
        }

        // 6. Send the Personalized Email
        MailApp.sendEmail(recipientEmail, personalizedSubject, "", {htmlBody: personalizedBody, attachments: attachments});
        Logger.log(`Email sent successfully to ${recipientEmail} for row ${currentRowNumber}.`);
        sentCount++;

      } catch (rowError) {
        // Catch and log errors specific to processing a single row
        const errorMessage = `Failed to send email for row ${currentRowNumber} (Recipient: ${row[emailColumnIndex] || 'N/A'}): ${rowError.message}`;
        Logger.log(errorMessage);
        errors.push(errorMessage);
        failedCount++;
      }
    });

    // 7. Provide Overall Feedback to the User
    let finalMessage = `Email sending complete!\n\nSent: ${sentCount}\nFailed: ${failedCount}`;
    if (errors.length > 0) {
      finalMessage += '\n\nSome emails failed to send or had attachment issues. Check "View > Executions" in the Apps Script editor for detailed logs.';
      ui.alert('Sending Complete with Errors', finalMessage, ui.ButtonSet.OK);
    } else {
      ui.alert('Sending Complete', finalMessage, ui.ButtonSet.OK);
    }

  } catch (mainError) {
    // Catch any critical errors that prevent the entire process from starting
    Logger.log(`Critical error during email sending process: ${mainError.message}`);
    ui.alert('Error Sending Emails', `A critical error occurred: ${mainError.message}. Please check the script logs for more details.`, ui.ButtonSet.OK);
  }
}