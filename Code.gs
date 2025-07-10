/**
 * @file Code.gs
 * @description Google Apps Script functions for the custom text editor.
 * This script handles the creation of the custom menu and display of the editor dialog.
 * It also handles saving and loading content to/from the Google Sheet.
 */

const TEMPLATE_SHEET_NAME = 'TextEditorTemplate'; // A more descriptive and unique name
const DOC_LINK_HEADER_NAME = 'DocLink'; // Standard header name for the Google Doc ID/URL column

// NEW: Constants for template sheet cell locations
const SUBJECT_CELL = 'B1';
const BODY_CELL = 'B2';

/**
 * Creates a custom menu in the Google Sheet when the spreadsheet is opened.
 * This function is triggered automatically when the spreadsheet loads.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Tools')
      .addItem('Open Email Composer', 'showTextEditorDialog')
      .addSeparator()
      .addItem('Send Personalized Emails', 'sendPersonalizedEmails')
      .addToUi();
}

/**
 * Displays the custom text editor in a modal dialog within the Google Sheet.
 * The HTML content for the editor is loaded from 'index.html'.
 */
function showTextEditorDialog() {
  const htmlOutput = HtmlService.createTemplateFromFile('index.html')
      .evaluate()
      .setWidth(850)
      .setHeight(600)
      .setTitle('Email Composer');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, htmlOutput.getTitle());
}

/**
 * Helper function to include external HTML, CSS, or JavaScript files
 * into a main HTML template file using Apps Script scriplets.
 *
 * @param {string} filename The name of the Apps Script HTML file to include.
 * @returns {string} The content of the specified file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Helper function to get a sheet by name, creating it if it doesn't exist.
 * If created, the sheet will be immediately hidden and resized.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The active spreadsheet.
 * @param {string} sheetName The name of the sheet to get or create.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Sheet object.
 * @throws {Error} If the sheet cannot be accessed or created.
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    try {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.hideSheet();

      const maxRows = sheet.getMaxRows();
      const maxColumns = sheet.getMaxColumns();

      if (maxColumns > 4) {
        sheet.deleteColumns(5, maxColumns - 4);
      }
      if (maxRows > 10) {
        sheet.deleteRows(11, maxRows - 10);
      }

      Logger.log('Sheet "' + sheetName + '" created, hidden, and resized to 10 rows and 4 columns.');
    } catch (e) {
      throw new Error('Could not create sheet "' + sheetName + '". Please check permissions or sheet name: ' + e.message);
    }
  }
  return sheet;
}

/**
 * NEW NAME: Saves the provided subject and HTML body content to the specified template sheet.
 * This function is called from the client-side JavaScript using google.script.run.
 *
 * @param {string} subject The subject line template to save.
 * @param {string} body The HTML string content (body) to save.
 */
function saveTemplate(subject, body) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(spreadsheet, TEMPLATE_SHEET_NAME);

    // Save subject to B1 and body to B2
    sheet.getRange(SUBJECT_CELL).setValue(subject);
    sheet.getRange(BODY_CELL).setValue(body);

    Logger.log('Template (Subject & Body) saved successfully to ' + TEMPLATE_SHEET_NAME + '!' + SUBJECT_CELL + ' and ' + BODY_CELL);
  } catch (e) {
    Logger.log('Error saving template: ' + e.message);
    throw new Error('Failed to save template: ' + e.message);
  }
}

/**
 * NEW NAME: Loads the subject and body content from the specified template sheet.
 * This function is called from the client-side JavaScript using google.script.run.
 *
 * @returns {Object} An object containing the subject and body properties.
 */
function loadTemplate() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(spreadsheet, TEMPLATE_SHEET_NAME);

    // Load subject from B1 and body from B2
    const subject = sheet.getRange(SUBJECT_CELL).getValue();
    const body = sheet.getRange(BODY_CELL).getValue();

    Logger.log('Template (Subject & Body) loaded successfully from ' + TEMPLATE_SHEET_NAME + '!' + SUBJECT_CELL + ' and ' + BODY_CELL);
    return { subject: subject.toString(), body: body.toString() }; // Ensure they are strings
  } catch (e) {
    Logger.log('Error loading template: ' + e.message);
    throw new Error('Failed to load template: ' + e.message);
  }
}

/**
 * Retrieves the header row (first row) from the active spreadsheet sheet.
 * @returns {string[]} An array of header names.
 */
function getSpreadsheetHeaders() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet(); // Get the currently active sheet (where the user's data is)
    if (!sheet) {
      throw new Error("No active sheet found in the spreadsheet.");
    }

    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      return []; // No columns, no headers
    }

    // Get all values from the first row
    const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
    Logger.log('Headers retrieved: ' + JSON.stringify(headers));
    return headers;
  } catch (e) {
    Logger.log('Error getting spreadsheet headers: ' + e.message);
    throw new Error('Failed to get spreadsheet headers: ' + e.message);
  }
}

/**
 * Helper function for basic email address validation.
 * @param {string} email The email string to validate.
 * @returns {boolean} True if the email is likely valid, false otherwise.
 */
function validateEmail(email) {
  const emailRegex = /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return emailRegex.test(String(email).toLowerCase());
}

/**
 * Extracts a Google Doc ID from a URL or returns the ID if it's already one.
 * @param {string} docLink A Google Doc URL or a direct Google Doc ID.
 * @returns {string|null} The extracted Google Doc ID, or null if not found.
 */
function extractDocId(docLink) {
  if (!docLink) return null;

  const urlRegex = /document\/d\/([a-zA-Z0-9_-]+)/;
  const match = docLink.match(urlRegex);

  if (match && match[1]) {
    return match[1];
  } else if (docLink.length > 20 && !docLink.includes('/')) {
    return docLink;
  }
  return null;
}

/**
 * Sends personalized emails based on the template and spreadsheet data, including document attachments.
 * This function is triggered via a custom menu item in the Google Sheet.
 */
function sendPersonalizedEmails() {
  const ui = SpreadsheetApp.getUi();

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = spreadsheet.getActiveSheet();

    // 1. Get Email Template (Subject & Body)
    const templateSheet = getOrCreateSheet(spreadsheet, TEMPLATE_SHEET_NAME);
    const emailSubjectTemplate = templateSheet.getRange(SUBJECT_CELL).getValue(); // NEW: Get subject template
    const emailBodyTemplate = templateSheet.getRange(BODY_CELL).getValue(); // Renamed from emailTemplate

    if (!emailSubjectTemplate && !emailBodyTemplate) { // If both are empty
      ui.alert('Email Template Missing', 'Both the subject and body templates are empty in the "' + TEMPLATE_SHEET_NAME + '" sheet. Please create them using the Text Editor.', ui.ButtonSet.OK);
      return;
    }
    // If only body is empty, we can still proceed with subject or vice versa.

    // 2. Get Spreadsheet Data (using getDisplayValues for formatting)
    const allData = dataSheet.getDataRange().getDisplayValues();

    if (allData.length < 2) {
      ui.alert('No Data Found', 'The active sheet must contain at least a header row and one row of data to send emails.', ui.ButtonSet.OK);
      return;
    }

    const headers = allData[0];
    const dataRows = allData.slice(1);

    // 3. Find Required Column Indices
    const emailColumnIndex = headers.findIndex(header => header.toLowerCase().trim() === 'email');
    const docLinkColumnIndex = headers.findIndex(header => header.toLowerCase().trim() === DOC_LINK_HEADER_NAME.toLowerCase());

    if (emailColumnIndex === -1) {
      ui.alert('Missing Email Column', 'The active sheet must have a column with the header "Email" to send emails.', ui.ButtonSet.OK);
      return;
    }

    let sentCount = 0;
    let failedCount = 0;
    const errors = [];

    // 4. Loop Through Rows, Replace Field Codes, and Send Email
    dataRows.forEach((row, rowIndex) => {
      const currentRowNumber = rowIndex + 2;
      try {
        const recipientEmail = row[emailColumnIndex];

        if (!recipientEmail || !validateEmail(recipientEmail)) {
          const errorMessage = `Invalid or missing email address in row ${currentRowNumber}: '${recipientEmail}'. Skipping this row.`;
          Logger.log(errorMessage);
          errors.push(errorMessage);
          failedCount++;
          return;
        }

        // Apply field code replacement to Subject
        let personalizedSubject = emailSubjectTemplate;
        // Apply field code replacement to Body
        let personalizedBody = emailBodyTemplate;

        headers.forEach((header, colIndex) => {
          const fieldValue = row[colIndex] !== undefined && row[colIndex] !== null ? String(row[colIndex]) : '';
          const fieldCode = `{{${header}}}`;
          const regex = new RegExp(fieldCode.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');

          // Replace in subject
          if (personalizedSubject) {
              personalizedSubject = personalizedSubject.replace(regex, fieldValue);
          }
          // Replace in body
          if (personalizedBody) {
              personalizedBody = personalizedBody.replace(regex, fieldValue);
          }
        });

        // 5. Handle Document Attachment
        let attachments = [];
        if (docLinkColumnIndex !== -1) {
          const docLinkValue = row[docLinkColumnIndex];
          if (docLinkValue) {
            const docId = extractDocId(docLinkValue);
            if (docId) {
              try {
                const file = DriveApp.getFileById(docId);
                const pdfBlob = file.getAs(MimeType.PDF);
                attachments.push(pdfBlob);
                Logger.log(`Attached ${file.getName()} (as PDF) for row ${currentRowNumber}.`);
              } catch (fileError) {
                const errorMessage = `Could not attach document for row ${currentRowNumber} (Link: ${docLinkValue}): ${fileError.message}. Skipping attachment for this email.`;
                Logger.log(errorMessage);
                errors.push(errorMessage);
              }
            } else {
              const errorMessage = `Invalid or unparseable document link in row ${currentRowNumber}: '${docLinkValue}'. Skipping attachment for this email.`;
              Logger.log(errorMessage);
              errors.push(errorMessage);
            }
          }
        }

        // 6. Send Email with Attachments
        MailApp.sendEmail(recipientEmail, personalizedSubject, "", {htmlBody: personalizedBody, attachments: attachments}); // Use personalizedSubject
        Logger.log(`Email sent successfully to ${recipientEmail} for row ${currentRowNumber}.`);
        sentCount++;

      } catch (rowError) {
        const errorMessage = `Failed to send email for row ${currentRowNumber} (Recipient: ${row[emailColumnIndex] || 'N/A'}): ${rowError.message}`;
        Logger.log(errorMessage);
        errors.push(errorMessage);
        failedCount++;
      }
    });

    // Provide overall feedback to the user
    let finalMessage = `Email sending complete!\nSent: ${sentCount}\nFailed: ${failedCount}`;
    if (errors.length > 0) {
      finalMessage += '\n\nSome emails failed to send or had attachment issues. Check "View > Executions" in the Apps Script editor for details.';
      ui.alert('Sending Complete with Errors', finalMessage, ui.ButtonSet.OK);
    } else {
      ui.alert('Sending Complete', finalMessage, ui.ButtonSet.OK);
    }

  } catch (mainError) {
    Logger.log('Critical error during email sending: ' + mainError.message);
    ui.alert('Error Sending Emails', 'A critical error occurred: ' + mainError.message + '. Please check the script logs for more details.', ui.ButtonSet.OK);
  }
}