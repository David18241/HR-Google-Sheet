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
  
  try {
    // 1. Extract and validate employee data
    const requiredFields = ['firstName', 'lastName', 'workEmail', 'startDate', 'jobFolderId', 'jobClassification'];
    const employeeData = extractAndValidateEmployeeData(PERSONNEL_SHEET_NAME, requiredFields);
    if (!employeeData) return;

    // 2. Create folder structure
    const folders = createEmployeeFolders(employeeData);
    if (!folders) return;

    // 3. Generate documents
    const documents = generateEmployeeDocuments(employeeData, folders);
    if (!documents) return;

    // 4. Set permissions
    if (!setEmployeePermissions(employeeData, folders, documents)) return;

    // 5. Update records
    if (!updateEmployeeRecords(employeeData, folders)) return;

    // 6. Add to groups
    if (!addEmployeeToGroups(employeeData)) return;

    // 7. Send notifications
    if (!sendEmployeeNotifications(employeeData, documents)) return;

    // --- Success ---
    closeSidebar();
    ui.alert(`Onboarding process completed successfully for ${employeeData.fullName}. Check drafts for Name Tag and Welcome emails.`);
    Logger.log(`Onboarding successful for ${employeeData.employeeFolderName} (Row ${employeeData.rowIndex}).`);

  } catch (error) {
    closeSidebar();
    handleWorkflowError(error, 'Onboarding workflow', 'Onboarding process failed. Please check the logs and manually complete any remaining steps.');
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
  try {
    const result = processAttestationEmails(templateId, emailSubject, attestationType);
    
    if (result) {
      const summaryMessage = `${attestationType} Attestation Email Summary:\n` +
                           `Successfully sent: ${result.emailsSent}\n` +
                           `Skipped or failed: ${result.emailsFailed}`;
      SpreadsheetApp.getUi().alert(summaryMessage);
      Logger.log(summaryMessage);
    }
  } catch (error) {
    closeSidebar();
    handleWorkflowError(error, `${attestationType} attestation workflow`, `Failed to send ${attestationType} attestation emails`);
  }
}

/**
 * Processes attestation emails for all active employees.
 * 
 * @param {string} templateId Template ID for the email
 * @param {string} emailSubject Email subject line
 * @param {string} attestationType Type of attestation (OSHA, HIPAA, etc.)
 * @returns {Object|null} Results object with counts or null on failure
 */
function processAttestationEmails(templateId, emailSubject, attestationType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get sheet data
  const sheetInfo = getSheetDataWithHeaders(PERSONNEL_SHEET_NAME, ss);
  if (!sheetInfo) return null;
  const { data, headerMap } = sheetInfo;

  // Validate required headers
  const requiredHeaders = [COL_FIRST_NAME, COL_WORK_EMAIL, COL_ACTIVE];
  const validation = validateRequiredFields(headerMap, requiredHeaders, `${attestationType} attestation headers`);
  if (!validation.isValid) {
    SpreadsheetApp.getUi().alert(`Error: Missing required columns in "${PERSONNEL_SHEET_NAME}" sheet:\n${validation.missingFields.join(', ')}`);
    return null;
  }

  const firstNameIndex = headerMap[COL_FIRST_NAME];
  const emailIndex = headerMap[COL_WORK_EMAIL];
  const isActiveIndex = headerMap[COL_ACTIVE];

  // Process emails
  let emailsSent = 0;
  let emailsFailed = 0;
  const totalEmployees = data.length - 1; // Exclude header row

  showProgressSidebar(`${attestationType} Attestations`, 'Initializing...', 0, totalEmployees);

  for (let i = 1; i < data.length; i++) { // Start at 1 to skip header
    const row = data[i];
    const firstName = row[firstNameIndex];
    const email = row[emailIndex];
    const isActive = String(row[isActiveIndex]).trim().toLowerCase();

    // Update progress
    showProgressSidebar(`${attestationType} Attestations`, `Processing ${firstName || 'employee'}...`, i, totalEmployees);

    // Check if active and has required info
    if (isActive === 'yes' && firstName && email && isValidEmail(email)) {
      const placeholders = { "{{FirstName}}": firstName };
      const success = createEmailFromTemplate(
        templateId, 
        email, 
        emailSubject, 
        placeholders, 
        false, // Send directly
        `${attestationType} Attestation - ${firstName}`
      );
      
      if (success) {
        emailsSent++;
      } else {
        emailsFailed++;
      }
      
      // Small delay to avoid quota issues
      Utilities.sleep(300);
      
    } else if (isActive === 'yes' && (!firstName || !email || !isValidEmail(email))) {
      Logger.log(`Skipping ${attestationType} email for row ${i + 1}: Missing or invalid data for active employee.`);
      emailsFailed++;
    }
    // Else: Not active, skip silently
  }

  closeSidebar();
  
  return { emailsSent, emailsFailed };
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
  try {
    // Extract and validate employee data
    const requiredFields = ['firstName', 'personalEmail'];
    const employeeData = extractAndValidateEmployeeData(PERSONNEL_SHEET_NAME, requiredFields);
    if (!employeeData) return;

    // Create draft
    const placeholders = { "{{FirstName}}": employeeData.firstName };
    const success = createEmailFromTemplate(
      BEFORE_FIRST_DAY_EMAIL_ID,
      employeeData.personalEmail,
      'Altamonte Dermatology: Getting Ready for Your First Day!',
      placeholders,
      true, // Create draft
      `Before First Day - ${employeeData.firstName}`
    );

    if (success) {
      SpreadsheetApp.getUi().alert(`"Before First Day" email draft created successfully for ${employeeData.firstName}. Please check your Gmail drafts.`);
      Logger.log(`Before First Day email draft created for ${employeeData.employeeFolderName}.`);
    }
  } catch (error) {
    handleWorkflowError(error, 'Before First Day email workflow', 'Failed to create Before First Day email draft');
  }
}


/**
 * Workflow for offboarding an employee based on the selected row.
 */
function offboardEmployeeWorkflow() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 1. Extract and validate employee data
    const requiredFields = ['firstName', 'lastName', 'workEmail'];
    const employeeData = extractAndValidateEmployeeData(PERSONNEL_SHEET_NAME, requiredFields);
    if (!employeeData) return;

    // 2. Confirmation dialog
    const confirmationMessage = `Are you sure you want to initiate the offboarding process for ${employeeData.fullName} (Row ${employeeData.rowIndex})?\n\n` +
                               `This will:\n` +
                               `- Mark employee as inactive\n` +
                               `- Remove from Google Groups\n` +
                               `- Archive folders and remove access\n` +
                               `- Create offboarding documentation\n\n` +
                               `This action affects active systems and should only be done when authorized.`;
    
    if (!showConfirmationDialog('Confirm Offboarding', confirmationMessage, 'OFFBOARD')) {
      ui.alert('Offboarding cancelled.');
      Logger.log('Offboarding cancelled by user.');
      return;
    }

    // If no end date provided, use today
    if (!employeeData.endDate) {
      employeeData.endDate = new Date();
      employeeData.formattedEndDate = formatDate(employeeData.endDate);
    }

    // 3. Mark as inactive
    showProgressSidebar('Offboarding', `Processing offboarding for ${employeeData.employeeFolderName}...`, 1, 5);
    if (!markEmployeeInactive(employeeData)) return;

    // 4. Remove from groups  
    showProgressSidebar('Offboarding', `Removing from groups...`, 2, 5);
    if (!removeEmployeeFromGroups(employeeData)) return;

    // 5. Archive folders and remove access
    showProgressSidebar('Offboarding', `Archiving folders...`, 3, 5);
    if (!archiveEmployeeFolders(employeeData)) return;

    // 6. Create offboarding documentation
    showProgressSidebar('Offboarding', `Creating documentation...`, 4, 5);
    createOffboardingDocumentation(employeeData);

    // 7. Create offboarding email (optional)
    showProgressSidebar('Offboarding', `Creating email notifications...`, 5, 5);
    createOffboardingEmailNotification(employeeData);

    // --- Success ---
    closeSidebar();
    ui.alert(`Offboarding process completed for ${employeeData.fullName}.\n\nPlease complete any additional manual steps:\n- System access reviews\n- Equipment collection\n- Final documentation review`);
    Logger.log(`Offboarding completed for ${employeeData.employeeFolderName} (Row ${employeeData.rowIndex}).`);

  } catch (error) {
    closeSidebar();
    handleWorkflowError(error, 'Offboarding workflow', 'Offboarding process failed. Please check the logs and manually complete required steps.');
  }
}

/**
 * Creates offboarding documentation for an employee.
 *
 * @param {Object} employeeData Employee data object
 * @returns {boolean} True if successful, false otherwise
 */
function createOffboardingDocumentation(employeeData) {
  try {
    // Only create documentation if template is available
    if (!OFFBOARDING_TEMPLATE_ID) {
      Logger.log('No offboarding template ID configured. Skipping document creation.');
      return true;
    }

    // TODO: Determine destination folder - for now, use the job folder
    const destinationFolder = employeeData.jobFolderId ? DriveApp.getFolderById(employeeData.jobFolderId) : null;
    
    if (!destinationFolder) {
      Logger.log('No destination folder available for offboarding document. Skipping document creation.');
      return true;
    }

    const offboardingPlaceholders = {
      "{{Employee Name}}": employeeData.fullName,
      "{{First Name}}": employeeData.firstName,
      "{{Last Name}}": employeeData.lastName,
      "{{End Date}}": employeeData.formattedEndDate,
      "{{Job Classification}}": employeeData.jobClassification || 'N/A'
    };

    const offboardingDoc = copyAndPersonalizeDocument(
      OFFBOARDING_TEMPLATE_ID,
      destinationFolder,
      `${employeeData.employeeFolderName} - Offboarding Checklist`,
      offboardingPlaceholders
    );

    if (offboardingDoc) {
      Logger.log(`Offboarding checklist created for ${employeeData.employeeFolderName}: ${offboardingDoc.getUrl()}`);
    }

    return true;
  } catch (error) {
    return handleWorkflowError(error, 'Offboarding documentation', `Failed to create offboarding documentation for ${employeeData.employeeFolderName}`, false);
  }
}

/**
 * Creates offboarding email notification.
 *
 * @param {Object} employeeData Employee data object
 * @returns {boolean} True if successful, false otherwise
 */
function createOffboardingEmailNotification(employeeData) {
  try {
    // Only create email if template is available
    if (!OFFBOARDING_EMAIL_ID) {
      Logger.log('No offboarding email template ID configured. Skipping email creation.');
      return true;
    }

    const offboardingEmailPlaceholders = {
      "{{FirstName}}": employeeData.firstName,
      "{{LastName}}": employeeData.lastName,
      "{{End Date}}": employeeData.formattedEndDate
    };

    // Create draft email to the employee
    const emailSuccess = createEmailFromTemplate(
      OFFBOARDING_EMAIL_ID,
      employeeData.workEmail,
      'Important Information Regarding Your Departure',
      offboardingEmailPlaceholders,
      true, // Create as draft
      `Offboarding Email - ${employeeData.employeeFolderName}`
    );

    if (emailSuccess) {
      Logger.log(`Offboarding email draft created for ${employeeData.employeeFolderName}.`);
    }

    return true;
  } catch (error) {
    return handleWorkflowError(error, 'Offboarding email', `Failed to create offboarding email for ${employeeData.employeeFolderName}`, false);
  }
}