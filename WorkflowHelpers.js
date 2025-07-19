/**
 * @fileoverview Helper functions for breaking down large workflow operations.
 */

/**
 * Extracts and validates employee data from the active row.
 *
 * @param {string} sheetName Name of the sheet to read from
 * @param {Array<string>} requiredFields Array of required field names
 * @returns {Object|null} Employee data object or null if validation fails
 */
function extractAndValidateEmployeeData(sheetName, requiredFields) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get Sheet Data and Headers
  const sheetInfo = getSheetDataWithHeaders(sheetName, ss);
  if (!sheetInfo) return null;
  const { headerMap, sheet } = sheetInfo;

  // Get Active Row Data
  const activeRowInfo = getActiveRowData(sheetName, headerMap, ss);
  if (!activeRowInfo) return null;
  const { rowData, rowIndex, range } = activeRowInfo;

  // Create employee data object
  const employeeData = {
    firstName: rowData[headerMap[COL_FIRST_NAME]],
    lastName: rowData[headerMap[COL_LAST_NAME]],
    workEmail: rowData[headerMap[COL_WORK_EMAIL]],
    personalEmail: rowData[headerMap[COL_PERSONAL_EMAIL]],
    startDate: rowData[headerMap[COL_START_DATE]] ? new Date(rowData[headerMap[COL_START_DATE]]) : null,
    endDate: rowData[headerMap[COL_END_DATE]] ? new Date(rowData[headerMap[COL_END_DATE]]) : null,
    jobFolderId: rowData[headerMap[COL_JOB_FOLDER_ID]],
    medRecFolderId: MEDREC_FOLDER_ID,
    jobClassification: rowData[headerMap[COL_JOB_CLASSIFICATION]],
    groupEmail: rowData[headerMap[COL_GROUP_EMAIL]],
    isActive: rowData[headerMap[COL_ACTIVE]],
    employeeFolderId: headerMap[COL_EMPLOYEE_DRIVE_FOLDER_ID] ? rowData[headerMap[COL_EMPLOYEE_DRIVE_FOLDER_ID]] : null,
    employeeMedRecFolderId: headerMap[COL_EMPLOYEE_MEDREC_FOLDER_ID] ? rowData[headerMap[COL_EMPLOYEE_MEDREC_FOLDER_ID]] : null,
    rowIndex: rowIndex,
    sheet: sheet,
    headerMap: headerMap
  };

  // Add computed fields
  employeeData.fullName = `${employeeData.firstName} ${employeeData.lastName}`;
  employeeData.employeeFolderName = `${employeeData.lastName}, ${employeeData.firstName}`;
  employeeData.formattedStartDate = employeeData.startDate ? formatDate(employeeData.startDate) : '';
  employeeData.formattedEndDate = employeeData.endDate ? formatDate(employeeData.endDate) : '';

  // Validate required fields
  const validation = validateRequiredFields(employeeData, requiredFields, `Employee data validation for row ${rowIndex}`);
  if (!validation.isValid) {
    SpreadsheetApp.getUi().alert(
      `Missing essential information in row ${rowIndex}:\n${validation.missingFields.join(', ')}\n\nPlease complete all required fields.`
    );
    return null;
  }

  // Validate email formats
  if (employeeData.workEmail && !isValidEmail(employeeData.workEmail)) {
    SpreadsheetApp.getUi().alert(`Invalid work email format in row ${rowIndex}: ${employeeData.workEmail}`);
    return null;
  }

  if (employeeData.personalEmail && !isValidEmail(employeeData.personalEmail)) {
    SpreadsheetApp.getUi().alert(`Invalid personal email format in row ${rowIndex}: ${employeeData.personalEmail}`);
    return null;
  }

  // Validate start date if provided
  if (employeeData.startDate) {
    const dateValidation = validateDate(employeeData.startDate, true);
    if (!dateValidation.isValid) {
      SpreadsheetApp.getUi().alert(`Invalid start date in row ${rowIndex}: ${dateValidation.message}`);
      return null;
    }
  }

  return employeeData;
}

/**
 * Creates the necessary folder structure for an employee.
 *
 * @param {Object} employeeData Employee data object
 * @returns {Object|null} Object containing created folders or null on failure
 */
function createEmployeeFolders(employeeData) {
  try {
    showProgressSidebar('Creating Folders', `Creating folders for ${employeeData.employeeFolderName}...`);
    
    // Create Medical Record folder
    const employeeMedRecFolder = createFolderIfNotExists(employeeData.employeeFolderName, MEDREC_FOLDER_ID);
    if (!employeeMedRecFolder) {
      throw new Error(`Failed to create Medical Record folder for ${employeeData.employeeFolderName}.`);
    }

    // Create main HR folder
    const employeeFolder = createFolderIfNotExists(employeeData.employeeFolderName, employeeData.jobFolderId);
    if (!employeeFolder) {
      throw new Error(`Failed to create main HR folder for ${employeeData.employeeFolderName}.`);
    }

    // Create Employee Access subfolder
    const employeeAccessFolderName = `${employeeData.employeeFolderName} ${EMPLOYEE_ACCESS_FOLDER_SUFFIX}`;
    const employeeAccessFolder = createFolderIfNotExists(employeeAccessFolderName, employeeFolder.getId());
    if (!employeeAccessFolder) {
      throw new Error(`Failed to create Employee Access subfolder for ${employeeData.employeeFolderName}.`);
    }

    return {
      employeeMedRecFolder,
      employeeFolder,
      employeeAccessFolder
    };
  } catch (error) {
    return handleWorkflowError(error, 'Folder creation', `Failed to create folders for ${employeeData.employeeFolderName}`);
  }
}

/**
 * Generates personalized documents for an employee.
 *
 * @param {Object} employeeData Employee data object
 * @param {Object} folders Object containing created folders
 * @returns {Object|null} Object containing created documents or null on failure
 */
function generateEmployeeDocuments(employeeData, folders) {
  try {
    showProgressSidebar('Creating Documents', `Generating documents for ${employeeData.employeeFolderName}...`);
    var currentDate = new Date();
    var year31 = currentDate.getFullYear() + 31;

    // Create Hepatitis B Vaccination Form
    const hepBPlaceholders = {
      "{{Employee Name}}": employeeData.fullName,
      "{{First Name}}": employeeData.firstName,
      "{{Last Name}}": employeeData.lastName,
      "{{Date}}": employeeData.formattedStartDate,
      "{{31Years}}": year31
    };
    
    const hepBDoc = copyAndPersonalizeDocument(
      HEPB_VAX_FORM_ID, 
      folders.employeeMedRecFolder, 
      `${employeeData.employeeFolderName} - Hepatitis B Vaccination Form`, 
      hepBPlaceholders
    );
    
    if (!hepBDoc) {
      throw new Error(`Failed to create Hepatitis B document for ${employeeData.employeeFolderName}.`);
    }

    // Create Onboarding Checklist
    const onboardingDoc = personalizeOnboardingDocumentSpecial(
      folders.employeeFolder, 
      employeeData.employeeFolderName, 
      employeeData.formattedStartDate, 
      hepBDoc, 
      folders.employeeMedRecFolder, 
      "https://link.to.your.hipaa.policy.or.form", // TODO: Update with actual links
      "https://link.to.your.handbook.or.form"
    );
    
    if (!onboardingDoc) {
      throw new Error(`Failed to create Onboarding document for ${employeeData.employeeFolderName}.`);
    }

    return {
      hepBDoc,
      onboardingDoc
    };
  } catch (error) {
    return handleWorkflowError(error, 'Document generation', `Failed to generate documents for ${employeeData.employeeFolderName}`);
  }
}

/**
 * Sets appropriate permissions for employee files and folders.
 *
 * @param {Object} employeeData Employee data object
 * @param {Object} folders Object containing created folders
 * @param {Object} documents Object containing created documents
 * @returns {boolean} True if successful, false otherwise
 */
function setEmployeePermissions(employeeData, folders, documents) {
  try {
    showProgressSidebar('Setting Permissions', `Setting permissions for ${employeeData.employeeFolderName}...`);
    
    // Employee needs to sign Hep B form
    addEditorToFile(documents.hepBDoc, employeeData.workEmail);
    
    // Employee views and comments on content in their access folder
    addCommenterToFolder(folders.employeeAccessFolder, employeeData.workEmail);
    
    Logger.log(`Permissions set successfully for ${employeeData.employeeFolderName}.`);
    return true;
  } catch (error) {
    return handleWorkflowError(error, 'Permission setting', `Failed to set permissions for ${employeeData.employeeFolderName}`);
  }
}

/**
 * Updates spreadsheet records with new folder IDs and other data.
 *
 * @param {Object} employeeData Employee data object
 * @param {Object} folders Object containing created folders
 * @returns {boolean} True if successful, false otherwise
 */
function updateEmployeeRecords(employeeData, folders) {
  try {
    showProgressSidebar('Updating Records', `Updating records for ${employeeData.employeeFolderName}...`);
    
    // Update spreadsheet with folder IDs
    if (employeeData.headerMap.hasOwnProperty(COL_EMPLOYEE_DRIVE_FOLDER_ID)) {
      employeeData.sheet.getRange(employeeData.rowIndex, employeeData.headerMap[COL_EMPLOYEE_DRIVE_FOLDER_ID] + 1)
        .setValue(folders.employeeFolder.getId());
    } else {
      Logger.log(`Column "${COL_EMPLOYEE_DRIVE_FOLDER_ID}" not found. Cannot write employee folder ID.`);
    }

    if (employeeData.headerMap.hasOwnProperty(COL_EMPLOYEE_MEDREC_FOLDER_ID)) {
      employeeData.sheet.getRange(employeeData.rowIndex, employeeData.headerMap[COL_EMPLOYEE_MEDREC_FOLDER_ID] + 1)
        .setValue(folders.employeeMedRecFolder.getId());
    } else {
      Logger.log(`Column "${COL_EMPLOYEE_MEDREC_FOLDER_ID}" not found. Cannot write med rec folder ID.`);
    }

    // Update Access Log
    updateAccessLog(
      ACCESS_LOG_SS_ID, 
      ACCESS_LOG_SHEET_NAME, 
      employeeData.employeeFolderName, 
      employeeData.formattedStartDate, 
      employeeData.jobClassification
    );

    Logger.log(`Records updated successfully for ${employeeData.employeeFolderName}.`);
    return true;
  } catch (error) {
    return handleWorkflowError(error, 'Record updating', `Failed to update records for ${employeeData.employeeFolderName}`);
  }
}

/**
 * Adds employee to appropriate Google Groups.
 *
 * @param {Object} employeeData Employee data object
 * @returns {boolean} True if successful, false otherwise
 */
function addEmployeeToGroups(employeeData) {
  try {
    showProgressSidebar('Adding to Groups', `Adding ${employeeData.employeeFolderName} to groups...`);
    
    addMembershipToGroups(employeeData.workEmail, employeeData.groupEmail, PRACTICE_GROUP_EMAIL);
    
    Logger.log(`Group memberships added successfully for ${employeeData.employeeFolderName}.`);
    return true;
  } catch (error) {
    return handleWorkflowError(error, 'Group membership', `Failed to add ${employeeData.employeeFolderName} to groups`);
  }
}

/**
 * Creates and sends notification emails for new employee.
 *
 * @param {Object} employeeData Employee data object
 * @param {Object} documents Object containing created documents
 * @returns {boolean} True if successful, false otherwise
 */
function sendEmployeeNotifications(employeeData, documents) {
  try {
    showProgressSidebar('Sending Notifications', `Creating email notifications for ${employeeData.employeeFolderName}...`);
    
    // Name Tag Email
    const nameTagPlaceholders = { "{{FirstName}}": employeeData.firstName };
    const nameTagSuccess = createEmailFromTemplate(
      NAMETAG_EMAIL_ID, 
      NAMETAG_VENDOR_EMAIL, 
      `Additional Name Tag for ${employeeData.fullName}`, 
      nameTagPlaceholders, 
      true, 
      `Name Tag Request - ${employeeData.employeeFolderName}`
    );

    // Welcome Email
    const welcomePlaceholders = {
      "{{FirstName}}": employeeData.firstName,
      "{{hepBDoc}}": { text: "Hepatitis B Vaccination Form", url: documents.hepBDoc.getUrl() }
    };
    const welcomeSuccess = createEmailFromTemplate(
      WELCOME_EMAIL_ID, 
      employeeData.workEmail, 
      'Welcome to the Team!', 
      welcomePlaceholders, 
      true, 
      `Welcome Email - ${employeeData.employeeFolderName}`
    );

    if (nameTagSuccess && welcomeSuccess) {
      Logger.log(`Email notifications created successfully for ${employeeData.employeeFolderName}.`);
      return true;
    } else {
      Logger.log(`Some email notifications failed for ${employeeData.employeeFolderName}.`);
      return false;
    }
  } catch (error) {
    return handleWorkflowError(error, 'Email notifications', `Failed to create email notifications for ${employeeData.employeeFolderName}`);
  }
}

/**
 * Marks an employee as inactive and sets end date.
 *
 * @param {Object} employeeData Employee data object
 * @returns {boolean} True if successful, false otherwise
 */
function markEmployeeInactive(employeeData) {
  try {
    // Mark as Inactive
    if (employeeData.headerMap.hasOwnProperty(COL_ACTIVE)) {
      employeeData.sheet.getRange(employeeData.rowIndex, employeeData.headerMap[COL_ACTIVE] + 1).setValue('No');
    }
    
    // Set End Date (use provided end date or today)
    if (employeeData.headerMap.hasOwnProperty(COL_END_DATE)) {
      const endDate = employeeData.endDate || new Date();
      employeeData.sheet.getRange(employeeData.rowIndex, employeeData.headerMap[COL_END_DATE] + 1).setValue(endDate);
    }

    Logger.log(`Employee ${employeeData.employeeFolderName} marked as inactive.`);
    return true;
  } catch (error) {
    return handleWorkflowError(error, 'Employee deactivation', `Failed to mark ${employeeData.employeeFolderName} as inactive`);
  }
}

/**
 * Removes employee from Google Groups.
 *
 * @param {Object} employeeData Employee data object
 * @returns {boolean} True if successful, false otherwise
 */
function removeEmployeeFromGroups(employeeData) {
  try {
    removeMembershipFromGroups(employeeData.workEmail, employeeData.groupEmail, PRACTICE_GROUP_EMAIL);
    
    Logger.log(`Group memberships removed successfully for ${employeeData.employeeFolderName}.`);
    return true;
  } catch (error) {
    return handleWorkflowError(error, 'Group removal', `Failed to remove ${employeeData.employeeFolderName} from groups`);
  }
}

/**
 * Archives employee folders and removes their access.
 *
 * @param {Object} employeeData Employee data object
 * @returns {boolean} True if successful, false otherwise
 */
function archiveEmployeeFolders(employeeData) {
  try {
    // This is a placeholder for folder archiving logic
    // TODO: Implement folder archiving based on your requirements
    
    if (employeeData.employeeFolderId) {
      const employeeFolder = DriveApp.getFolderById(employeeData.employeeFolderId);
      // Remove employee access
      try {
        employeeFolder.removeViewer(employeeData.workEmail);
        employeeFolder.removeEditor(employeeData.workEmail);
      } catch (removeError) {
        Logger.log(`Could not remove access for ${employeeData.workEmail}: ${removeError.message}`);
      }
      
      // Move to Former Employees folder (check if it's in a shared drive first)
      if (FORMER_EMPLOYEES_FOLDER_ID) {
        try {
          const formerEmployeesFolder = DriveApp.getFolderById(FORMER_EMPLOYEES_FOLDER_ID);
          const originalParentFolder = employeeFolder.getParents().next();
          
          // Try to move the folder - if it fails due to shared drive, catch the error
          formerEmployeesFolder.addFolder(employeeFolder);
          originalParentFolder.removeFolder(employeeFolder);
          Logger.log(`Employee folder moved to Former Employees folder for ${employeeData.employeeFolderName}`);
        } catch (moveError) {
          if (moveError.message.includes("shared drive")) {
            Logger.log(`Employee folder for ${employeeData.employeeFolderName} is in a shared drive and cannot be moved. Access has been removed instead.`);
          } else {
            Logger.log(`Could not move employee folder for ${employeeData.employeeFolderName}: ${moveError.message}`);
          }
        }
      }
      
      Logger.log(`Folder access removed for ${employeeData.employeeFolderName}`);
    } else {
      Logger.log(`Could not archive HR folder: Employee Folder ID not found for ${employeeData.employeeFolderName}.`);
    }

    // Similar process for medical records folder
    if (employeeData.employeeMedRecFolderId) {
      const medRecFolder = DriveApp.getFolderById(employeeData.employeeMedRecFolderId);
      try {
        medRecFolder.removeViewer(employeeData.workEmail);
        medRecFolder.removeEditor(employeeData.workEmail);
      } catch (removeError) {
        Logger.log(`Could not remove med rec access for ${employeeData.workEmail}: ${removeError.message}`);
      }
    }

    return true;
  } catch (error) {
    return handleWorkflowError(error, 'Folder archiving', `Failed to archive folders for ${employeeData.employeeFolderName}`);
  }
}