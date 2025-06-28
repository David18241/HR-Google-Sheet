/**
 * @fileoverview Main functions for HR processes, including menu setup and core actions.
 */

/**
 * Creates the custom menu in the spreadsheet UI when opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Employee Actions')
      .addItem('Onboard New Employee (row highlighted)', 'onboardEmployeeWorkflow') // Renamed for clarity
      .addItem('Send OSHA Training Attestation to All Active Employees', 'sendOshaAttestationWorkflow')
      .addItem('Send HIPAA Training Attestation to All Active Employees', 'sendHipaaAttestationWorkflow')
      .addItem('Generate Before First Day Email Draft (row highlighted)', 'createBeforeFirstDayEmailWorkflow')
      .addItem('Offboard Employee (row highlighted)', 'offboardEmployeeWorkflow')
      .addSeparator()
      .addItem('Check Skills for Renewal (Provider Training Sheet)', 'checkSkillsForEmployeeRenewals') // Moved from Utilities for menu access
      .addToUi();
}

/**
 * Workflow for onboarding a new employee based on the selected row.
 */
function onboardEmployeeWorkflow() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get Sheet Data and Headers
  const sheetInfo = getSheetDataWithHeaders(PERSONNEL_SHEET_NAME, ss);
  if (!sheetInfo) return; // Error handled in helper
  const { headerMap } = sheetInfo;

  // 2. Get Active Row Data
  const activeRowInfo = getActiveRowData(PERSONNEL_SHEET_NAME, headerMap, ss);
  if (!activeRowInfo) return; // Error handled in helper
  const { rowData, rowIndex, range } = activeRowInfo;

  // 3. Extract Employee Information
  const firstName = rowData[headerMap[COL_FIRST_NAME]];
  const lastName = rowData[headerMap[COL_LAST_NAME]];
  const employeeEmail = rowData[headerMap[COL_WORK_EMAIL]];
  const startDateObj = rowData[headerMap[COL_START_DATE]] ? new Date(rowData[headerMap[COL_START_DATE]]) : null;
  const jobFolderId = rowData[headerMap[COL_JOB_FOLDER_ID]];
  const medRecFolderId = rowData[headerMap[COL_MED_REC_FOLDER_ID]];
  const job = rowData[headerMap[COL_JOB_CLASSIFICATION]];
  const groupEmail = rowData[headerMap[COL_GROUP_EMAIL]];

  // Validate essential info
  if (!firstName || !lastName || !employeeEmail || !startDateObj || isNaN(startDateObj.getTime()) || !jobFolderId || !medRecFolderId || !job) {
      ui.alert(`Missing essential information in the selected row (Row ${rowIndex}). Please ensure First Name, Last Name, Work Email, Start Date, Job Folder ID, Med Rec Folder ID, and Primary Classification are filled correctly.`);
      Logger.log(`Onboarding aborted for row ${rowIndex} due to missing data.`);
      return;
  }

  const formattedStartDate = formatDate(startDateObj);
  const employeeFolderName = `${lastName}, ${firstName}`;
  const employeeAccessFolderName = `${employeeFolderName} ${EMPLOYEE_ACCESS_FOLDER_SUFFIX}`;

  // --- Start Process ---
  ui.showSidebar(HtmlService.createHtmlOutput('<p>Processing onboarding for ' + employeeFolderName + '... Please wait.</p>').setWidth(300).setHeight(100));
  SpreadsheetApp.flush(); // Ensure UI updates


  try {
      // 4. Create Folders
      const employeeMedRecFolder = createFolderIfNotExists(employeeFolderName, medRecFolderId);
      if (!employeeMedRecFolder) throw new Error(`Failed to create Medical Record folder for ${employeeFolderName}.`);

      const employeeFolder = createFolderIfNotExists(employeeFolderName, jobFolderId);
      if (!employeeFolder) throw new Error(`Failed to create main HR folder for ${employeeFolderName}.`);

      const employeeAccessFolder = createFolderIfNotExists(employeeAccessFolderName, employeeFolder.getId());
       if (!employeeAccessFolder) throw new Error(`Failed to create Employee Access subfolder for ${employeeFolderName}.`);


      // 5. Create and Personalize Documents
      const hepBPlaceholders = {
          "{{Employee Name}}": `${firstName} ${lastName}`,
          "{{First Name}}": firstName,
          "{{Last Name}}": lastName,
          "{{Date}}": formattedStartDate
      };
      const hepBDoc = copyAndPersonalizeDocument(HEPB_VAX_FORM_ID, employeeMedRecFolder, `${employeeFolderName} - Hepatitis B Vaccination Form`, hepBPlaceholders);
      if (!hepBDoc) throw new Error(`Failed to create Hepatitis B document for ${employeeFolderName}.`);

      // Deprecated Docs (kept comments from original)
      // const hipaaDoc = personalizeHipaaAgreementDocument(employeeFolder, employeeFolderName, formattedStartDate); // Now handled by form
      // const handbookDoc = personalizeHandbookDocument(employeeFolder, employeeFolderName, formattedStartDate); // Assuming handbook is separate or linked in onboarding doc
      // Using placeholders for where these *would* go if they were still generated docs
      const placeholderHipaaLink = "https://link.to.your.hipaa.policy.or.form"; // Replace with actual link
      const placeholderHandbookLink = "https://link.to.your.handbook.or.form"; // Replace with actual link


      const onboardingPlaceholders = {
          "{{Employee Name}}": `${firstName} ${lastName}`,
          "{{First Name}}": firstName,
          "{{Last Name}}": lastName,
          "{{Start Date}}": formattedStartDate,
           // Placeholders for links - using the link replacement feature in the email helper now
         // "{{hipaaDoc}}": { text: "Confidentiality and Security Agreement", url: placeholderHipaaLink }, //Deprecated after transition to google form
         // "{{handbookDoc}}": { text: "Handbook Acknowledgement", url: placeholderHandbookLink }, //Deprecated after transition to google form
          "{{hepBDoc}}": { text: "Hepatitis B Vaccination Form", url: hepBDoc.getUrl() }, // Use actual Hep B doc URL
          "{{medRecFolder}}": { text: "Medical Record Folder", url: employeeMedRecFolder.getUrl()}
      };
      // Need a specialized function for onboarding doc due to link replacements within list items
      // For now, using the old function structure, needs refactor to use generic helpers better
      const onboardingDoc = personalizeOnboardingDocumentSpecial(employeeFolder, employeeFolderName, formattedStartDate, hepBDoc, employeeMedRecFolder, placeholderHipaaLink, placeholderHandbookLink);
      if (!onboardingDoc) throw new Error(`Failed to create Onboarding document for ${employeeFolderName}.`);


      // 6. Set Permissions
      addEditorToFile(hepBDoc, employeeEmail); // Employee needs to sign Hep B form
      addCommenterToFolder(employeeAccessFolder, employeeEmail); // Employee views and comments on content in their access folder
       // Decide if employee needs edit access to the main onboarding checklist itself
       // addEditorToFile(onboardingDoc, employeeEmail);


       // 7. Update Spreadsheet with Folder IDs (RECOMMENDED)
       // Check if columns exist, if not, add them? Or just log a warning.
        if (headerMap.hasOwnProperty(COL_EMPLOYEE_DRIVE_FOLDER_ID)) {
             sheetInfo.sheet.getRange(rowIndex, headerMap[COL_EMPLOYEE_DRIVE_FOLDER_ID] + 1).setValue(employeeFolder.getId());
        } else Logger.log(`Column "${COL_EMPLOYEE_DRIVE_FOLDER_ID}" not found. Cannot write employee folder ID.`);

        if (headerMap.hasOwnProperty(COL_EMPLOYEE_MEDREC_FOLDER_ID)) {
             sheetInfo.sheet.getRange(rowIndex, headerMap[COL_EMPLOYEE_MEDREC_FOLDER_ID] + 1).setValue(employeeMedRecFolder.getId());
        } else Logger.log(`Column "${COL_EMPLOYEE_MEDREC_FOLDER_ID}" not found. Cannot write med rec folder ID.`);


      // 8. Update Access Log
      updateAccessLog(ACCESS_LOG_SS_ID, ACCESS_LOG_SHEET_NAME, employeeFolderName, formattedStartDate, job);

      // 9. Add to Google Groups
      addMembershipToGroups(employeeEmail, groupEmail, PRACTICE_GROUP_EMAIL);

      // 10. Create Email Drafts
      // Name Tag Email
      const nameTagPlaceholders = { "{{FirstName}}": firstName };
      createEmailFromTemplate(NAMETAG_EMAIL_ID, NAMETAG_VENDOR_EMAIL, `Additional Name Tag for ${firstName} ${lastName}`, nameTagPlaceholders, true, `Name Tag Request - ${employeeFolderName}`);

      // Welcome Email
      const welcomePlaceholders = {
        "{{FirstName}}": firstName,
        "{{hepBDoc}}": { text: "Hepatitis B Vaccination Form", url: hepBDoc.getUrl() } // Pass link object
      };
      createEmailFromTemplate(WELCOME_EMAIL_ID, employeeEmail, 'Welcome to the Team!', welcomePlaceholders, true, `Welcome Email - ${employeeFolderName}`);

      // --- Finish Process ---
      ui.alert(`Onboarding process completed successfully for ${firstName} ${lastName}. Check drafts for Name Tag and Welcome emails.`);
      Logger.log(`Onboarding successful for ${employeeFolderName} (Row ${rowIndex}).`);

  } catch (error) {
      Logger.log(`ONBOARDING FAILED for ${employeeFolderName} (Row ${rowIndex}): ${error.message} \nStack: ${error.stack}`);
      ui.alert(`Onboarding process FAILED for ${firstName} ${lastName}. Error: ${error.message}. Please check the logs for details and manually complete any remaining steps.`);
  }
}

// Temporary function until personalizeOnboardingDocument is fully refactored
// to use replacePlaceholderWithLink correctly within list items
function personalizeOnboardingDocumentSpecial(employeeFolder, employeeName, startDate, hepBDoc, medRecFldr, hipaaUrl, handbookUrl) {
  const ui = SpreadsheetApp.getUi();
  try {
      const template = DriveApp.getFileById(ONBOARDING_TEMPLATE_ID);
      const nameParts = parseEmployeeName(employeeName); // Use helper
      const firstName = nameParts.firstName;
      const lastName = nameParts.lastName;
      const fullName = nameParts.fullName;

      const newFileName = `${employeeName} - Onboarding Form`;
      const newDocFile = template.makeCopy(newFileName, employeeFolder);
      const newDoc = DocumentApp.openById(newDocFile.getId());
      const body = newDoc.getBody();

      // Replace simple placeholders
      body.replaceText("{{Employee Name}}", fullName);
      body.replaceText("{{First Name}}", firstName);
      body.replaceText("{{Last Name}}", lastName);
      body.replaceText("{{Start Date}}", startDate);

      // --- Link Replacements in List Items (Needs careful handling) ---
      // This part is tricky with the generic helper because findText might not work well across list item boundaries
      // Keeping the original list item loop logic for now.
      // const placeholderHipaaText = "{{hipaaDoc}}";  //Deprecated after transition to google form
      // const replacementHipaaText = "Confidentiality and Security Agreement";  //Deprecated after transition to google form
      // const placeholderHandbookText = "{{handbookDoc}}";  //Deprecated after transition to google form
      // const replacementHandbookText = "Handbook Acknowledgement";  //Deprecated after transition to google form
      const placeholderHepBText = "{{hepBDoc}}";
      const replacementHepBText = "Hepatitis B Vaccination Form";
      const placeholderMedRecFldrText = "{{medRecFolder}}";
      const replacementMedRecFldrText = "Medical Record Folder";

      const listItems = body.getListItems();
      for (const listItem of listItems) {
          const textElement = listItem.editAsText(); // Get Text object from ListItem
          // Find and replace HIPAA placeholder
          // replaceListItemPlaceholderWithLink(textElement, placeholderHipaaText, replacementHipaaText, hipaaUrl); //Deprecated after transition to google form
          // Find and replace Handbook placeholder
          // replaceListItemPlaceholderWithLink(textElement, placeholderHandbookText, replacementHandbookText, handbookUrl);  //Deprecated after transition to google form
           // Find and replace Hep B placeholder
          replaceListItemPlaceholderWithLink(textElement, placeholderHepBText, replacementHepBText, hepBDoc.getUrl());
          // Find and replace Med Rec Folder placeholder
          replaceListItemPlaceholderWithLink(textElement, placeholderMedRecFldrText, replacementMedRecFldrText, medRecFldr.getUrl());
      }

      newDoc.saveAndClose();
      Logger.log(`Created and personalized onboarding document for ${fullName}. URL: ${newDocFile.getUrl()}`);
      return newDoc; // Return Document object

  } catch (e) {
      Logger.log(`Error personalizing onboarding document for ${employeeName}: ${e.message} \nStack: ${e.stack}`);
      ui.alert(`Error creating the Onboarding checklist document for ${employeeName}. Check Template ID and logs.`);
      // Clean up potentially created file?
      return null;
  }
}

// Helper specifically for replacing placeholders within list item text elements
function replaceListItemPlaceholderWithLink(textElement, placeholder, linkText, url) {
    let foundRange;
    while (foundRange = textElement.findText(placeholder)) {
        const start = foundRange.getStartOffset();
        const end = foundRange.getEndOffsetInclusive();
        try {
             // Check if start/end are valid before proceeding
            if (start === -1 || end === -1 || start > end) {
                Logger.log(`Invalid range found for placeholder "${placeholder}" in list item. Skipping.`);
                break; // Exit loop for this placeholder in this item
            }
            textElement.deleteText(start, end);
            textElement.insertText(start, linkText);
            textElement.setLinkUrl(start, start + linkText.length - 1, url);
            // Continue searching from the beginning of the element in case of multiple occurrences
        } catch (linkError) {
            Logger.log(`Error applying link for "${placeholder}" -> "${linkText}" at offset ${start}: ${linkError.message}`);
            // Break to avoid potential infinite loops if error persists
            break;
        }
    }
}



/**
 * Workflow to send a specific training attestation email (OSHA or HIPAA) to all active employees.
 *
 * @param {string} templateId The Google Doc template ID for the email content.
 * @param {string} emailSubject The subject line for the email.
 * @param {string} attestationType A descriptive name (e.g., "OSHA", "HIPAA") for logging/alerts.
 */
function sendTrainingAttestationWorkflow(templateId, emailSubject, attestationType) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get Sheet Data
  const sheetInfo = getSheetDataWithHeaders(PERSONNEL_SHEET_NAME, ss);
  if (!sheetInfo) return;
  const { data, headerMap } = sheetInfo;

  // 2. Validate Headers
  const requiredHeaders = [COL_FIRST_NAME, COL_WORK_EMAIL, COL_ACTIVE];
  for (const header of requiredHeaders) {
    if (!headerMap.hasOwnProperty(header)) {
      ui.alert(`Error: Missing required column "${header}" in the "${PERSONNEL_SHEET_NAME}" sheet.`);
      Logger.log(`Missing header "${header}" for ${attestationType} attestation email.`);
      return;
    }
  }
  const firstNameIndex = headerMap[COL_FIRST_NAME];
  const emailIndex = headerMap[COL_WORK_EMAIL];
  const isActiveIndex = headerMap[COL_ACTIVE];

  // 3. Loop through employees and send emails
  let emailsSent = 0;
  let emailsFailed = 0;
  const totalEmployees = data.length - 1; // Exclude header row

  ui.showSidebar(HtmlService.createHtmlOutput(`<p>Sending ${attestationType} attestations... </p><p id="status">Processing 0/${totalEmployees}</p>`).setWidth(300).setHeight(100), `${attestationType} Sending Status`);
  SpreadsheetApp.flush();


  for (let i = 1; i < data.length; i++) { // Start at 1 to skip header
    const row = data[i];
    const firstName = row[firstNameIndex];
    const email = row[emailIndex];
    const isActive = String(row[isActiveIndex]).trim().toLowerCase(); // Normalize 'Yes'/'No'

     // Update status in sidebar (optional, can slow down if many employees)
    // Use eval to run client-side script in sidebar
    const statusUpdateScript = `<script>document.getElementById('status').innerText = 'Processing ${i}/${totalEmployees}';</script>`;
    ui.showSidebar(HtmlService.createHtmlOutput(`<p>Sending ${attestationType} attestations... </p><p id="status">Processing ${i}/${totalEmployees}</p>${statusUpdateScript}`).setWidth(300).setHeight(100), `${attestationType} Sending Status`);
    SpreadsheetApp.flush();


    // Check if active and has required info
    if (isActive === 'yes' && firstName && email) {
      const placeholders = { "{{FirstName}}": firstName };
      const success = createEmailFromTemplate(templateId, email, emailSubject, placeholders, false, `${attestationType} Attestation - ${firstName}`); // false = send directly
      if (success) {
        emailsSent++;
      } else {
        emailsFailed++;
        // Error already logged and alerted in createEmailFromTemplate
      }
        // Add a small delay to avoid exceeding quotas if sending many emails
        Utilities.sleep(500); // Sleep for 500 milliseconds

    } else if (isActive === 'yes' && (!firstName || !email)) {
        Logger.log(`Skipping ${attestationType} email for row ${i + 1}: Missing First Name or Work Email for active employee.`);
        emailsFailed++; // Count as failed/skipped
    }
     // Else: Not active, skip silently
  }

  ui.closeSidebar();

  // 4. Final Report
  let summaryMessage = `${attestationType} Attestation Email Summary:\n`;
  summaryMessage += `Successfully sent: ${emailsSent}\n`;
  summaryMessage += `Skipped or failed: ${emailsFailed}`;
  ui.alert(summaryMessage);
  Logger.log(summaryMessage);
}

/** Trigger for sending OSHA attestation emails. */
function sendOshaAttestationWorkflow() {
  sendTrainingAttestationWorkflow(
    OSHA_TRAINING_ATTESTATION_EMAIL_TEMPLATE_ID,
    'OSHA Training Attestation Required',
    'OSHA'
  );
}

/** Trigger for sending HIPAA attestation emails. */
function sendHipaaAttestationWorkflow() {
  sendTrainingAttestationWorkflow(
    HIPAA_TRAINING_ATTESTATION_EMAIL_TEMPLATE_ID,
    'HIPAA Training Attestation Required',
    'HIPAA'
  );
}

/**
 * Workflow to create the "Before First Day" email draft for the selected employee.
 */
function createBeforeFirstDayEmailWorkflow() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get Sheet Data and Headers
  const sheetInfo = getSheetDataWithHeaders(PERSONNEL_SHEET_NAME, ss);
  if (!sheetInfo) return;
  const { headerMap } = sheetInfo;

  // 2. Get Active Row Data
  const activeRowInfo = getActiveRowData(PERSONNEL_SHEET_NAME, headerMap, ss);
  if (!activeRowInfo) return;
  const { rowData, rowIndex } = activeRowInfo;

  // 3. Extract Required Info
  const firstName = rowData[headerMap[COL_FIRST_NAME]];
  const personalEmail = rowData[headerMap[COL_PERSONAL_EMAIL]]; // Use Personal Email

  // Validate
  if (!firstName || !personalEmail) {
      ui.alert(`Missing First Name or Personal Email in the selected row (Row ${rowIndex}). Cannot create draft.`);
      Logger.log(`Before First Day email draft creation aborted for row ${rowIndex} due to missing data.`);
      return;
  }

  // 4. Create Draft
  const placeholders = { "{{FirstName}}": firstName };
  const success = createEmailFromTemplate(
      BEFORE_FIRST_DAY_EMAIL_ID,
      personalEmail,
      'Altamonte Dermatology: Getting Ready for Your First Day!', // Updated subject
      placeholders,
      true, // true = create draft
      `Before First Day - ${firstName}`
  );

  if (success) {
      ui.alert(`"Before First Day" email draft created successfully for ${firstName}. Please check your Gmail drafts.`);
  }
  // Failure message handled by createEmailFromTemplate
}


/**
 * Workflow for offboarding an employee based on the selected row.
 * (Placeholder - implement actual offboarding steps)
 */
function offboardEmployeeWorkflow() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get Sheet Data and Headers
  const sheetInfo = getSheetDataWithHeaders(PERSONNEL_SHEET_NAME, ss);
  if (!sheetInfo) return;
  const { headerMap, sheet } = sheetInfo;

  // 2. Get Active Row Data
  const activeRowInfo = getActiveRowData(PERSONNEL_SHEET_NAME, headerMap, ss);
  if (!activeRowInfo) return;
  const { rowData, rowIndex, range } = activeRowInfo;

  // 3. Extract Employee Information
  const firstName = rowData[headerMap[COL_FIRST_NAME]];
  const lastName = rowData[headerMap[COL_LAST_NAME]];
  const employeeEmail = rowData[headerMap[COL_WORK_EMAIL]];
  const endDateObj = rowData[headerMap[COL_END_DATE]] ? new Date(rowData[headerMap[COL_END_DATE]]) : new Date(); // Default to today if empty
  const groupEmail = rowData[headerMap[COL_GROUP_EMAIL]];
  const employeeFolderName = `${lastName}, ${firstName}`; // Needed for finding folders/docs if IDs aren't stored

   // Optional: Get folder IDs if you stored them during onboarding
   const employeeFolderId = headerMap[COL_EMPLOYEE_DRIVE_FOLDER_ID] ? rowData[headerMap[COL_EMPLOYEE_DRIVE_FOLDER_ID]] : null;
   const employeeMedRecFolderId = headerMap[COL_EMPLOYEE_MEDREC_FOLDER_ID] ? rowData[headerMap[COL_EMPLOYEE_MEDREC_FOLDER_ID]] : null;


  // 4. Confirmation Dialog
  const confirmation = ui.prompt(
      'Confirm Offboarding',
      `Are you sure you want to initiate the offboarding process for ${firstName} ${lastName} (Row ${rowIndex})?\n\nThis will typically involve:\n- Marking as inactive\n- Removing from groups\n- Archiving folders (optional)\n- Creating offboarding checklist/email\n\nEnter 'OFFBOARD' to confirm:`,
      ui.ButtonSet.OK_CANCEL);

  if (confirmation.getSelectedButton() !== ui.Button.OK || confirmation.getResponseText().toUpperCase() !== 'OFFBOARD') {
      ui.alert('Offboarding cancelled.');
      Logger.log('Offboarding cancelled by user.');
      return;
  }

  // --- Start Process ---
  ui.showSidebar(HtmlService.createHtmlOutput('<p>Processing offboarding for ' + employeeFolderName + '... Please wait.</p>').setWidth(300).setHeight(100), 'Offboarding Status');
  SpreadsheetApp.flush();

  try {
      // TODO: Implement actual offboarding steps:

      // 5. Mark as Inactive & Set End Date in Sheet
      if (headerMap.hasOwnProperty(COL_ACTIVE)) {
          sheet.getRange(rowIndex, headerMap[COL_ACTIVE] + 1).setValue('No');
      }
       if (headerMap.hasOwnProperty(COL_END_DATE)) {
            // Format date before setting, or set as Date object
            sheet.getRange(rowIndex, headerMap[COL_END_DATE] + 1).setValue(endDateObj);
       }


      // 6. Remove from Google Groups
      removeMembershipFromGroups(employeeEmail, groupEmail, PRACTICE_GROUP_EMAIL);

      // 7. Handle Drive Access / Archive Folders
      //    - Remove editor/viewer access from employee-specific folders/files
      //    - Potentially move employee folders (HR and MedRec) to an 'Archived Employees' parent folder.
      //    - This requires finding the folders (ideally using stored IDs).
      // Example (needs refinement and error handling):
      /*
      if (employeeFolderId) {
          const employeeFolder = DriveApp.getFolderById(employeeFolderId);
          employeeFolder.removeViewer(employeeEmail); // Or removeEditor if they had it
          // Move to archive folder (get ARCHIVE_FOLDER_ID from Constants.gs)
          // DriveApp.getFolderById(ARCHIVE_FOLDER_ID).addFolder(employeeFolder);
          // DriveApp.getRootFolder().removeFolder(employeeFolder); // Remove from original parent *after* adding to archive
          Logger.log(`Removed access and archived HR folder for ${employeeFolderName}`);
      } else {
          Logger.log(`Could not archive HR folder automatically: Employee Folder ID not found in sheet for row ${rowIndex}.`);
      }
      // Repeat for MedRec folder
      */


      // 8. Create Offboarding Checklist Document (Similar to Onboarding)
      // const offboardingPlaceholders = { "{{Employee Name}}": `${firstName} ${lastName}`, "{{EndDate}}": formatDate(endDateObj) };
      // copyAndPersonalizeDocument(OFFBOARDING_TEMPLATE_ID, /* Destination Folder? Maybe Admin's folder? */, `${employeeFolderName} - Offboarding Checklist`, offboardingPlaceholders);


      // 9. Create Offboarding Email Draft (Optional - to employee or internal?)
      // const offboardingEmailPlaceholders = { "{{FirstName}}": firstName, "{{LastName}}": lastName, "{{EndDate}}": formatDate(endDateObj) };
      // createEmailFromTemplate(OFFBOARDING_EMAIL_ID, employeeEmail /* or internal HR email? */, 'Offboarding Information', offboardingEmailPlaceholders, true, `Offboarding Email - ${employeeFolderName}`);


      // --- Finish Process ---
       ui.closeSidebar();
      ui.alert(`Offboarding process initiated for ${firstName} ${lastName}. Please complete any manual steps (e.g., final checks, system deactivations).`);
      Logger.log(`Offboarding initiated for ${employeeFolderName} (Row ${rowIndex}).`);

  } catch (error) {
      ui.closeSidebar();
      Logger.log(`OFFBOARDING FAILED for ${employeeFolderName} (Row ${rowIndex}): ${error.message} \nStack: ${error.stack}`);
      ui.alert(`Offboarding process FAILED for ${firstName} ${lastName}. Error: ${error.message}. Please check the logs and manually complete required steps.`);
  }

}