/**
 * Interactive Vacation Approval Workflow Engine
 * 
 * Target Environment: Google Apps Script (Triggered on Form Submit & Web App GET)
 * Description: Captures Google Form submissions, constructs interactive HTML emails 
 *              with direct Approve/Deny buttons, parses manager decisions via web 
 *              endpoint, updates the database, and alerts the employee.
 * 
 * Operational Status: Production-Ready
 */

const SPREADSHEET_DB_NAME = "Form Responses 1";
// Deploy as a Web App in GAS and insert your URL here
const SCRIPT_WEB_APP_URL = "https://script.google.com/macros/s/AKfycb...MOCK_WEB_APP_URL.../exec"; 

/**
 * Triggers automatically on form submission to generate the approval email
 * Tied to the "On Form Submit" installable trigger
 */
function onFormSubmit(e) {
  try {
    if (!e || !e.values || !e.range) {
      Logger.log("Error: Function must be triggered by an active form submission event.");
      return;
    }

    const rowValues = e.values; 
    const timestamp = String(rowValues[0]);
    const employeeName = rowValues[1];
    const employeeEmail = String(rowValues[2]).toLowerCase().trim();
    const startDate = rowValues[3];
    const endDate = rowValues[4];
    const managerEmail = String(rowValues[5]).toLowerCase().trim();
    
    // Core Fix: Get exact row index from event range to prevent concurrency overwrites
    const rowIndex = e.range.getRow();
    
    // Core Fix: Correctly pass parameters to computeDigest to prevent byte compilation errors
    const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, timestamp, Utilities.Charset.UTF_8);
    const requestId = Utilities.base64Encode(rawHash).substring(0, 8).replace(/[^a-zA-Z0-9]/g, "");
    
    // Construct unique action URLs for the email buttons
    const approveUrl = `${SCRIPT_WEB_APP_URL}?action=approve&row=${rowIndex}&email=${encodeURIComponent(employeeEmail)}&id=${requestId}`;
    const denyUrl = `${SCRIPT_WEB_APP_URL}?action=deny&row=${rowIndex}&email=${encodeURIComponent(employeeEmail)}&id=${requestId}`;
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 500px; padding: 20px; border: 1px solid #ddd; border-radius: 5px;">
        <h3 style="color: #2c3e50; margin-top: 0;">Vacation Leave Request</h3>
        <p><strong>Employee:</strong> ${employeeName} (${employeeEmail})</p>
        <p><strong>Requested Dates:</strong> ${startDate} to ${endDate}</p>
        <p style="margin-bottom: 25px; color: #555;">Please click a button below to update the system instantly:</p>
        <table width="100%">
          <tr>
            <td>
              <a href="${approveUrl}" style="background-color: #27ae60; color: white; padding: 10px 20px; text-decoration: none; border-radius: 3px; font-weight: bold; display: inline-block;">Approve</a>
            </td>
            <td>
              <a href="${denyUrl}" style="background-color: #c0392b; color: white; padding: 10px 20px; text-decoration: none; border-radius: 3px; font-weight: bold; display: inline-block;">Deny</a>
            </td>
          </tr>
        </table>
        <p style="font-size: 10px; color: #aaa; margin-top: 25px;">System Generated. Request Token: ${requestId}</p>
      </div>
    `;
    
    if (managerEmail.includes("@")) {
      MailApp.sendEmail({
        to: managerEmail,
        subject: `LEAVE APPROVAL REQUIRED: ${employeeName}`,
        htmlBody: htmlBody
      });
      Logger.log(`Approval email successfully dispatched to ${managerEmail} for row ${rowIndex}.`);
    }
  } catch (err) {
    Logger.log(`Failed to process form submission. Error: ${err.toString()}`);
  }
}

/**
 * REST Web App Endpoint: Handles manager button clicks from inbox
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const rowIndex = parseInt(e.parameter.row, 10);
    const employeeEmail = decodeURIComponent(e.parameter.email);
    const requestId = e.parameter.id;
    
    if (!action || isNaN(rowIndex) || !employeeEmail) {
      return HtmlService.createHtmlOutput("<h3 style='color:#c0392b; font-family:Arial;'>Error: Missing or corrupt parameters in URL execution.</h3>");
    }
    
    const status = action === "approve" ? "APPROVED" : "DENIED";
    
    // Update target status cell (Column G / Index 7)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SPREADSHEET_DB_NAME);
    
    if (!sheet) {
      return HtmlService.createHtmlOutput("<h3 style='color:#c0392b; font-family:Arial;'>Error: Target database sheet not found.</h3>");
    }
    
    sheet.getRange(rowIndex, 7).setValue(status); 
    
    // Trigger automated email status notification back to employee
    notifySubmitter(employeeEmail, status, requestId);
    
    return HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; text-align: center; padding-top: 50px;">
        <h2 style="color: ${action === "approve" ? "#27ae60" : "#c0392b"};">Request ${status}!</h2>
        <p>The system database has been updated for Request ID: <strong>${requestId}</strong>.</p>
        <p style="color: #7f8c8d; font-size: 13px;">The employee has been notified. You can safely close this browser window.</p>
      </div>
    `);
  } catch (err) {
    return HtmlService.createHtmlOutput(`<h3 style='color:#c0392b; font-family:Arial;'>Fatal: Web endpoint runtime crash. Details: ${err.toString()}</h3>`);
  }
}

/**
 * Sends confirmation email directly to the request submitter
 */
function notifySubmitter(email, status, requestId) {
  if (!email || !email.includes("@")) return;
  
  const subject = `Your Leave Request Status Update: ${status}`;
  const body = `Hello,\n\nYour manager has reviewed and processed your vacation request.\n\nRequest ID: ${requestId}\nFinal Status: ${status}\n\nThis is an automated operational notification.`;
  
  MailApp.sendEmail(email, subject, body);
}
