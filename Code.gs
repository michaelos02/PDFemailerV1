/**
 * @file Code.gs
 * @description Google Apps Script functions for the Email Composer.
 * This script manages custom menu creation, dialog display,
 * template saving/loading, and personalized email sending with attachments.
 */

// --- Configuration Constants ---
const TEMPLATE_SHEET_NAME = 'EmailComposerTemplate'; // Name of the hidden sheet for storing email templates
const DOC_LINK_HEADER_NAME = 'DocLink'; // Standard header name for the Google Doc ID/URL column in the data sheet
const EMAIL_STATUS_HEADER_NAME = 'Email Status'; // Header for the email sending status column

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

  // Create the main Email Tools menu
  const emailToolsMenu = ui.createMenu('Email Tools')
      .addItem('Open Email Composer', 'showEmailComposerDialog')
      .addSeparator()
      .addItem('Send Personalized Emails', 'sendPersonalizedEmails');

  // Create the 'Help' submenu
  const helpSubMenu = ui.createMenu('Help')
      .addItem('User Guide', 'showUserGuideDialog')
      .addItem('Disclaimer', 'showDisclaimerDialog')
      .addItem('License', 'showLicenseDialog');

  // Add the 'Help' submenu to the main Email Tools menu
  emailToolsMenu.addSeparator() // Optional: Separator before the Help submenu
                .addSubMenu(helpSubMenu)
                .addToUi();
}

/**
 * Displays the custom Email Composer in a modal dialog within the Google Sheet.
 * The HTML content for the editor is loaded from 'index.html'.
 */
function showEmailComposerDialog() {
  const htmlOutput = HtmlService.createTemplateFromFile('index.html')
      .evaluate()
      .setWidth(850)
      .setHeight(600)
      .setTitle('Email Composer');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, htmlOutput.getTitle());
}

/**
 * Displays the user guide in a modal dialog.
 * The HTML content for the guide is loaded from 'userGuide.html'.
 */
function showUserGuideDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('userGuide.html')
      .setWidth(850) // Adjust width as needed
      .setHeight(600) // Adjust height as needed
      .setTitle('Personalized Email Sender: User Guide');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, htmlOutput.getTitle());
}

/**
 * Displays the disclaimer in a modal dialog.
 * The HTML content is loaded from 'disclaimer.html'.
 */
function showDisclaimerDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('disclaimer.html')
      .setWidth(600) // Adjust width as needed
      .setHeight(400) // Adjust height as needed
      .setTitle('Disclaimer');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, htmlOutput.getTitle());
}

/**
 * Displays the license information in a modal dialog.
 * The HTML content is loaded from 'license.html'.
 */
function showLicenseDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('license.html')
      .setWidth(700) // Adjust width as needed
      .setHeight(500) // Adjust height as needed
      .setTitle('License Information');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, htmlOutput.getTitle());
}
// --- Helper Functions ---

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

    Logger.log('Template (Subject & Body) saved successfully to ' + TEMPLATE_SHEET_NAME + '!' + SUBJECT_CELL + ' and ' + BODY_CELL + '.');
  } catch (e) {
    Logger.log('Error saving template: ' + e.message);
    throw new Error('Failed to save template: ' + e.message);
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

    Logger.log('Template (Subject & Body) loaded successfully from ' + TEMPLATE_SHEET_NAME + '!' + SUBJECT_CELL + ' and ' + BODY_CELL + '.');
    return { subject: subject.toString(), body: body.toString() }; // Ensure values are returned as strings
  } catch (e) {
    Logger.log('Error loading template: ' + e.message);
    throw new Error('Failed to load template: ' + e.message);
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

    // Get all values from the first row, respecting cell formatting (e.g., "$100.00", dates)
    const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
    Logger.log('Headers retrieved: ' + JSON.stringify(headers));
    return headers;
  } catch (e) {
    Logger.log('Error getting spreadsheet headers: ' + e.message);
    throw new Error('Failed to get spreadsheet headers: ' + e.message);
  }
}

// --- Email Sending Logic ---

/**
 * Original sendPersonalizedEmails function logic, now renamed to be called AFTER preview confirmation.
 * This function handles the actual batch sending of personalized emails.
 */
function executePersonalizedEmailSend() { // This function holds the core sending logic
  const ui = SpreadsheetApp.getUi();

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = spreadsheet.getActiveSheet();

    // 1. Retrieve Email Template (Subject and Body)
    const templateSheet = getOrCreateSheet(spreadsheet, TEMPLATE_SHEET_NAME);
    const emailSubjectTemplate = templateSheet.getRange(SUBJECT_CELL).getValue();
    const emailBodyTemplate = templateSheet.getRange(BODY_CELL).getValue();

    // 2. Get All Data from the Active Sheet
    const allData = dataSheet.getDataRange().getDisplayValues(); // Get display values for status check

    const headers = allData[0];
    const dataRows = allData.slice(1);
    const numRows = dataRows.length;

    // 3. Find Required Column Indices
    const emailColumnIndex = headers.findIndex(header => header.toLowerCase().trim() === 'email');
    const docLinkColumnIndex = headers.findIndex(header => header.toLowerCase().trim() === DOC_LINK_HEADER_NAME.toLowerCase());

    // 4. Determine Email Status Column and Add Header if Missing
    let statusColumnIndex = headers.findIndex(header => header.toLowerCase().trim() === EMAIL_STATUS_HEADER_NAME.toLowerCase());
    let currentLastColumn = dataSheet.getLastColumn();

    if (statusColumnIndex === -1) {
        statusColumnIndex = currentLastColumn;
        dataSheet.getRange(1, statusColumnIndex + 1).setValue(EMAIL_STATUS_HEADER_NAME);
        headers[statusColumnIndex] = EMAIL_STATUS_HEADER_NAME;
        Logger.log(`Added "${EMAIL_STATUS_HEADER_NAME}" header at column ${statusColumnIndex + 1}.`);
    } else {
        Logger.log(`"${EMAIL_STATUS_HEADER_NAME}" header found at column ${statusColumnIndex + 1}.`);
    }

    let sentCount = 0;
    let failedCount = 0;
    let skippedCount = 0; // NEW: Counter for skipped rows
    const errors = [];
    const statusMessages = new Array(numRows).fill('');
    const errorRowsToColor = [];

    // 5. Iterate Through Each Data Row to Personalize and Send Email
    dataRows.forEach((row, rowIndex) => {
      const currentRowNumber = rowIndex + 2; // Actual row number in sheet (1-indexed)

      // NEW: Check if Email Status is already populated
      const currentStatusInSheet = row[statusColumnIndex];
      if (currentStatusInSheet && currentStatusInSheet.toString().trim() !== '') {
          const skipMessage = `Skipped: Status already present ('${currentStatusInSheet}').`;
          Logger.log(`Row ${currentRowNumber}: ${skipMessage}`);
          statusMessages[rowIndex] = skipMessage;
          skippedCount++;
          return; // Skip to the next row
      }

      let rowStatus = ''; // Temporary status for the current row

      try {
        const recipientEmail = row[emailColumnIndex];

        // Basic validation for recipient email address
        if (!recipientEmail || !validateEmail(recipientEmail)) {
          rowStatus = `Error: Invalid or missing email ('${recipientEmail || 'empty'}').`;
          Logger.log(`Row ${currentRowNumber}: ${rowStatus}`);
          errors.push(`Row ${currentRowNumber}: ${rowStatus}`);
          failedCount++;
          errorRowsToColor.push(rowIndex);
          statusMessages[rowIndex] = rowStatus;
          return;
        }

        let personalizedSubject = emailSubjectTemplate;
        let personalizedBody = emailBodyTemplate;

        headers.forEach((header, colIndex) => {
          const fieldValue = row[colIndex] !== undefined && row[colIndex] !== null ? String(row[colIndex]) : '';
          const fieldCode = `[[${header}]]`;
          const regex = new RegExp(fieldCode.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');

          if (personalizedSubject) {
              personalizedSubject = personalizedSubject.replace(regex, fieldValue);
          }
          if (personalizedBody) {
              personalizedBody = personalizedBody.replace(regex, fieldValue);
          }
        });

        // 6. Handle Google Doc Attachment
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
                rowStatus += `Attachment Error: Could not attach document ('${docLinkValue}'): ${fileError.message}. `;
                Logger.log(`Row ${currentRowNumber}: ${rowStatus}`);
                errors.push(`Row ${currentRowNumber}: ${rowStatus}`);
                errorRowsToColor.push(rowIndex);
              }
            } else {
              rowStatus += `Attachment Error: Invalid or unparseable document link ('${docLinkValue}'). `;
              Logger.log(`Row ${currentRowNumber}: ${rowStatus}`);
              errors.push(`Row ${currentRowNumber}: ${rowStatus}`);
              errorRowsToColor.push(rowIndex);
            }
          }
        }

        // 7. Send Email with Attachments
        MailApp.sendEmail(recipientEmail, personalizedSubject, "", {htmlBody: personalizedBody, attachments: attachments});
        
        const sentTimestamp = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), "M/d/yy h:mm a");
        rowStatus = (rowStatus ? rowStatus + ' | ' : '') + `Email sent on ${sentTimestamp}.`;
        
        Logger.log(`Email sent successfully to ${recipientEmail} for row ${currentRowNumber}.`);
        sentCount++;

      } catch (rowError) {
        rowStatus = `Error: Failed to send email for '${row[emailColumnIndex] || 'N/A'}': ${rowError.message}.`;
        Logger.log(`Row ${currentRowNumber}: ${rowStatus}`);
        errors.push(`Row ${currentRowNumber}: ${rowStatus}`);
        failedCount++;
        errorRowsToColor.push(rowIndex);
      } finally {
        statusMessages[rowIndex] = rowStatus;
      }
    });

    // 8. Write all collected status messages back to the sheet in a single batch
    const statusValuesToWrite = statusMessages.map(msg => [msg]);
    if (numRows > 0) {
      const statusRange = dataSheet.getRange(2, statusColumnIndex + 1, numRows, 1);
      statusRange.setValues(statusValuesToWrite);

      // 9. Apply red color to error messages
      const fontColors = statusValuesToWrite.map((value, idx) => {
        return errorRowsToColor.includes(idx) ? 'red' : 'black';
      });
      statusRange.setFontColors(fontColors.map(color => [color]));

      Logger.log(`Status messages and colors written to column ${statusColumnIndex + 1}.`);
    }

    // 10. Provide overall feedback to the user (Updated Message)
    let finalMessage = `Email sending complete!\n\nSent: ${sentCount}\nFailed: ${failedCount}\nSkipped: ${skippedCount}`; // UPDATED: Added Skipped count
    if (errors.length > 0) {
      finalMessage += '\n\nSome emails failed to send or had attachment issues. Please check the "Email Status" column in your sheet for details.';
      ui.alert('Sending Complete with Errors', finalMessage, ui.ButtonSet.OK);
    } else {
      ui.alert('Sending Complete', finalMessage, ui.ButtonSet.OK);
    }

  } catch (mainError) {
    Logger.log('Critical error during email sending: ' + mainError.message);
    ui.alert('Error Sending Emails', 'A critical error occurred: ' + mainError.message + '. Please check the script logs for more details (or contact the script administrator if you are an end-user).', ui.ButtonSet.OK);
  }
}


/**
 * This function is triggered by the "Send Personalized Emails" menu item.
 * It prepares a preview of the first email and displays it in a dialog.
 * The actual sending is triggered by a button within the preview dialog.
 */
function sendPersonalizedEmails() { // This function is now the PREVIEW TRIGGER
  const ui = SpreadsheetApp.getUi();

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = spreadsheet.getActiveSheet();

    // 1. Retrieve Email Template (Subject and Body)
    const templateSheet = getOrCreateSheet(spreadsheet, TEMPLATE_SHEET_NAME);
    const emailSubjectTemplate = templateSheet.getRange(SUBJECT_CELL).getValue();
    const emailBodyTemplate = templateSheet.getRange(BODY_CELL).getValue();

    if (!emailSubjectTemplate && !emailBodyTemplate) {
      ui.alert('Email Template Missing', `Both the subject and body templates are empty in the "${TEMPLATE_SHEET_NAME}" sheet. Please create them using the Email Composer.`, ui.ButtonSet.OK);
      return;
    }

    // 2. Get All Data from the Active Sheet
    const allData = dataSheet.getDataRange().getDisplayValues();

    if (allData.length < 2) { // At least one header row and one data row are needed for preview
      ui.alert('No Data Found', 'The active sheet must contain at least a header row and one row of data to preview/send emails.', ui.ButtonSet.OK);
      return;
    }

    const headers = allData[0];
    const firstDataRow = allData[1]; // Get only the first data row for preview

    // 3. Find Required Column Indices (only email needed for basic check)
    const emailColumnIndex = headers.findIndex(header => header.toLowerCase().trim() === 'email');
    if (emailColumnIndex === -1) {
      ui.alert('Missing Email Column', 'The active sheet must have a column with the header "Email" to preview/send emails.', ui.ButtonSet.OK);
      return;
    }

    // 4. Personalize Subject and Body for the First Row
    let personalizedSubject = emailSubjectTemplate;
    let personalizedBody = emailBodyTemplate;

    headers.forEach((header, colIndex) => {
      const fieldValue = firstDataRow[colIndex] !== undefined && firstDataRow[colIndex] !== null ? String(firstDataRow[colIndex]) : '';
      const fieldCode = `[[${header}]]`;
      const regex = new RegExp(fieldCode.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');

      if (personalizedSubject) {
          personalizedSubject = personalizedSubject.replace(regex, fieldValue);
      }
      if (personalizedBody) {
          personalizedBody = personalizedBody.replace(regex, fieldValue);
      }
    });
    
    // 5. Open Preview Dialog and Pass Data
    const htmlTemplate = HtmlService.createTemplateFromFile('preview.html');
    htmlTemplate.personalizedSubject = personalizedSubject;
    htmlTemplate.personalizedBody = personalizedBody;

    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(700)
      .setHeight(500)
      .setTitle('Email Preview');

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, htmlOutput.getTitle());

  } catch (mainError) {
    Logger.log('Critical error during preview generation: ' + mainError.message);
    ui.alert('Error Generating Preview', 'A critical error occurred while preparing the preview: ' + mainError.message + '. Please check the script logs for more details.', ui.ButtonSet.OK);
  }
}