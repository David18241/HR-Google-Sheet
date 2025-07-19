/**
 * @fileoverview Helper functions using the Admin SDK Directory Service.
 *               Requires 'AdminDirectory' advanced service to be enabled.
 */

/**
 * Checks if a user is already a member of a Google Group.
 *
 * @param {string} userEmail The email address of the user to check.
 * @param {string} groupEmail The email address of the group to check.
 * @returns {boolean} True if user is a member, false if not or group doesn't exist.
 */
function isUserMemberOfGroup(userEmail, groupEmail) {
  if (!userEmail || !groupEmail) {
    return false;
  }
  
  try {
    const member = AdminDirectory.Members.get(groupEmail, userEmail);
    return member !== null;
  } catch (error) {
    if (error.message.includes("Resource Not Found") || error.message.includes("Member not found")) {
      return false; // User is not a member, or group doesn't exist
    }
    Logger.log(`Error checking membership for ${userEmail} in group ${groupEmail}: ${error.message}`);
    return false;
  }
}

/**
 * Checks if a Google Group exists.
 *
 * @param {string} groupEmail The email address of the group to check.
 * @returns {boolean} True if group exists, false otherwise.
 */
function doesGroupExist(groupEmail) {
  if (!groupEmail) {
    return false;
  }
  
  try {
    const group = AdminDirectory.Groups.get(groupEmail);
    return group !== null;
  } catch (error) {
    if (error.message.includes("Resource Not Found")) {
      return false; // Group doesn't exist
    }
    Logger.log(`Error checking if group exists ${groupEmail}: ${error.message}`);
    return false;
  }
}

/**
 * Sets the delivery settings for a group member to digest mode.
 *
 * @param {string} userEmail The email address of the user.
 * @param {string} groupEmail The email address of the group.
 * @returns {boolean} True if successful, false otherwise.
 */
function setMemberDeliveryToDigest(userEmail, groupEmail) {
  if (!userEmail || !groupEmail) {
    return false;
  }
  
  try {
    const deliverySettings = {
      'delivery_settings': 'digest'
    };
    
    AdminDirectory.Members.patch(deliverySettings, groupEmail, userEmail);
    Logger.log(`Set delivery to digest for ${userEmail} in group ${groupEmail}.`);
    return true;
  } catch (error) {
    Logger.log(`Error setting digest delivery for ${userEmail} in group ${groupEmail}: ${error.message}`);
    return false;
  }
}

/**
 * Adds a user email to specified Google Groups.
 * Adds to both a specific group (e.g., role-based) and the main practice group.
 * Handles "Member already exists" errors gracefully.
 *
 * @param {string} userEmail The email address of the user to add.
 * @param {string} specificGroupEmail The email address of the role-specific group. Can be null or empty.
 * @param {string} practiceWideGroupEmail The email address of the main practice group.
 */
function addMembershipToGroups(userEmail, specificGroupEmail, practiceWideGroupEmail) {
  if (!userEmail) {
    Logger.log("Skipping group membership: userEmail is missing.");
    return;
  }

  const member = {
    email: userEmail,
    role: 'MEMBER' // Or 'OWNER', 'MANAGER' if needed
  };

  // Add to specific group(s) if provided - handle comma-separated values
  if (specificGroupEmail) {
    if (!checkServiceAvailability("AdminDirectory", AdminDirectory)) {
      return;
    }
    
    // Parse comma-separated group emails from multi-select dropdown
    const groupEmails = specificGroupEmail.split(',').map(email => email.trim()).filter(email => email.length > 0);
    
    for (const groupEmail of groupEmails) {
      Logger.log(`Processing group: ${groupEmail}`);
      
      // Check if group exists first
      if (!doesGroupExist(groupEmail)) {
        Logger.log(`Group ${groupEmail} does not exist. Skipping addition of ${userEmail}.`);
      } else if (isUserMemberOfGroup(userEmail, groupEmail)) {
        Logger.log(`User ${userEmail} is already a member of group ${groupEmail}. Skipping addition.`);
      } else {
        // Group exists and user is not a member, proceed with addition
        const addToGroupOperation = () => {
          AdminDirectory.Members.insert(member, groupEmail);
          return true;
        };
        
        const result = executeWithRetry(
          addToGroupOperation, 
          `Adding ${userEmail} to group ${groupEmail}`
        );
        
        if (result) {
          Logger.log(`User ${userEmail} added as a member to the group ${groupEmail}.`);
          
          // Set delivery to digest for specific groups (not practice group)
          const digestResult = setMemberDeliveryToDigest(userEmail, groupEmail);
          if (digestResult) {
            Logger.log(`Delivery settings set to digest for ${userEmail} in group ${groupEmail}.`);
          }
        } else {
          Logger.log(`Failed to add ${userEmail} to group ${groupEmail} after retries.`);
        }
      }
    }
  } else {
      Logger.log(`Skipping adding ${userEmail} to specific group (no group email provided).`);
  }

  // Add to practice-wide group
  if (practiceWideGroupEmail) {
    if (!checkServiceAvailability("AdminDirectory", AdminDirectory)) {
      return;
    }
    
    // Check if group exists first
    if (!doesGroupExist(practiceWideGroupEmail)) {
      Logger.log(`Practice group ${practiceWideGroupEmail} does not exist. Skipping addition of ${userEmail}.`);
    } else if (isUserMemberOfGroup(userEmail, practiceWideGroupEmail)) {
      Logger.log(`User ${userEmail} is already a member of practice group ${practiceWideGroupEmail}. Skipping addition.`);
    } else {
      // Group exists and user is not a member, proceed with addition
      const addToPracticeGroupOperation = () => {
        AdminDirectory.Members.insert(member, practiceWideGroupEmail);
        return true;
      };
      
      const result = executeWithRetry(
        addToPracticeGroupOperation, 
        `Adding ${userEmail} to practice group ${practiceWideGroupEmail}`
      );
      
      if (result) {
        Logger.log(`User ${userEmail} added as a member to the group ${practiceWideGroupEmail}.`);
      } else {
        Logger.log(`Failed to add ${userEmail} to practice group ${practiceWideGroupEmail} after retries.`);
      }
    }
  } else {
      Logger.log(`Skipping adding ${userEmail} to practice-wide group (no group email provided).`);
  }
}

/**
 * Removes a user email from specified Google Groups.
 * Removes from both a specific group and the main practice group.
 * Handles "Resource Not Found: memberKey" errors gracefully.
 *
 * @param {string} userEmail The email address of the user to remove.
 * @param {string} specificGroupEmail The email address of the role-specific group. Can be null or empty.
 * @param {string} practiceWideGroupEmail The email address of the main practice group.
 */
function removeMembershipFromGroups(userEmail, specificGroupEmail, practiceWideGroupEmail) {
    if (!userEmail) {
        Logger.log("Skipping group membership removal: userEmail is missing.");
        return;
    }

     // Remove from specific group(s) if provided - handle comma-separated values
    if (specificGroupEmail) {
        if (!checkServiceAvailability("AdminDirectory", AdminDirectory)) {
            return;
        }
        
        // Parse comma-separated group emails from multi-select dropdown
        const groupEmails = specificGroupEmail.split(',').map(email => email.trim()).filter(email => email.length > 0);
        
        for (const groupEmail of groupEmails) {
            Logger.log(`Processing group removal: ${groupEmail}`);
            
            // Check if group exists first
            if (!doesGroupExist(groupEmail)) {
                Logger.log(`Group ${groupEmail} does not exist. Skipping removal of ${userEmail}.`);
            } else if (!isUserMemberOfGroup(userEmail, groupEmail)) {
                Logger.log(`User ${userEmail} is not a member of group ${groupEmail}. Skipping removal.`);
            } else {
                // Group exists and user is a member, proceed with removal
                const removeFromGroupOperation = () => {
                    AdminDirectory.Members.remove(groupEmail, userEmail);
                    return true;
                };
                
                const result = executeWithRetry(
                    removeFromGroupOperation, 
                    `Removing ${userEmail} from group ${groupEmail}`
                );
                
                if (result) {
                    Logger.log(`User ${userEmail} removed from the group ${groupEmail}.`);
                } else {
                    Logger.log(`Failed to remove ${userEmail} from group ${groupEmail} after retries.`);
                }
            }
        }
    } else {
      Logger.log(`Skipping removing ${userEmail} from specific group (no group email provided).`);
    }

    // Remove from practice-wide group
    if (practiceWideGroupEmail) {
        if (!checkServiceAvailability("AdminDirectory", AdminDirectory)) {
            return;
        }
        
        const removeFromPracticeGroupOperation = () => {
            AdminDirectory.Members.remove(practiceWideGroupEmail, userEmail);
            return true;
        };
        
        const result = executeWithRetry(
            removeFromPracticeGroupOperation, 
            `Removing ${userEmail} from practice group ${practiceWideGroupEmail}`
        );
        
        if (result) {
            Logger.log(`User ${userEmail} removed from the group ${practiceWideGroupEmail}.`);
        } else {
            Logger.log(`Failed to remove ${userEmail} from practice group ${practiceWideGroupEmail} after retries.`);
        }
    } else {
        Logger.log(`Skipping removing ${userEmail} from practice-wide group (no group email provided).`);
    }
}

/**
 * Gets default access values for a specific employee classification from the Default Access sheet.
 * Returns the checkbox values as boolean values for copying into the Employee sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} defaultAccessSheet The Default Access sheet.
 * @param {string} jobClassification The employee classification to look up.
 * @returns {Object|null} Object containing the default access values, or null if not found.
 */
function getDefaultAccessValues(defaultAccessSheet, jobClassification) {
    try {
        // Get all data from the Default Access sheet
        const dataRange = defaultAccessSheet.getDataRange();
        const values = dataRange.getValues();
        
        // Find the row with matching employee classification
        let matchingRowIndex = -1;
        for (let i = 1; i < values.length; i++) { // Start from row 2 (index 1)
            if (values[i][0] && values[i][0].toString().trim() === jobClassification.trim()) {
                matchingRowIndex = i;
                break;
            }
        }
        
        if (matchingRowIndex === -1) {
            Logger.log(`No matching classification found for "${jobClassification}" in Default Access sheet.`);
            return null;
        }
        
        const matchingRow = values[matchingRowIndex];
        
        // Column mapping for Default Access sheet based on the screenshot
        // A: Employee Classification, B: No Access, C: Health & financial info, etc.
        const defaultAccessValues = {
            noAccess: matchingRow[1] || false,                    // B - No Access
            healthFinancialInfo: matchingRow[2] || false,         // C - Health & financial information (EMR & PMS)
            emailCloudDrive: matchingRow[3] || false,             // D - Email and Cloud Drive
            textVoicemail: matchingRow[4] || false,               // E - Text and Voicemail
            voip: matchingRow[5] || false,                        // F - VOIP
            rems: matchingRow[6] || false                         // G - REMS
        };
        
        Logger.log(`Found default access values for "${jobClassification}": ${JSON.stringify(defaultAccessValues)}`);
        return defaultAccessValues;
        
    } catch (error) {
        Logger.log(`Error getting default access values for "${jobClassification}": ${error.message}`);
        return null;
    }
}

/**
 * Updates the Access Log spreadsheet with new employee information.
 * Creates a new row in the Employee sheet and populates it with default access values 
 * based on employee classification from the Default Access sheet.
 *
 * @param {string} spreadsheetId The ID of the Access Log spreadsheet.
 * @param {string} sheetName The name of the sheet within the Access Log (e.g., "Employees").
 * @param {string} employeeName Formatted employee name ("Last, First").
 * @param {string} startDate Formatted start date string.
 * @param {string} jobClassification The employee's job title/classification.
 */
function updateAccessLog(spreadsheetId, sheetName, employeeName, startDate, jobClassification) {
  try {
    const accessLogSs = SpreadsheetApp.openById(spreadsheetId);
    const employeeSheet = accessLogSs.getSheetByName(sheetName);
    const defaultAccessSheet = accessLogSs.getSheetByName("Default Access");

    if (!employeeSheet) {
        Logger.log(`Error: Access Log sheet "${sheetName}" not found in spreadsheet ID ${spreadsheetId}.`);
        SpreadsheetApp.getUi().alert(`Error: Could not find the sheet "${sheetName}" in the Access Log. Update skipped.`);
        return;
    }

    if (!defaultAccessSheet) {
        Logger.log(`Error: Default Access sheet not found in spreadsheet ID ${spreadsheetId}.`);
        SpreadsheetApp.getUi().alert(`Error: Could not find the "Default Access" sheet in the Access Log. Update skipped.`);
        return;
    }

    const lastRow = employeeSheet.getLastRow();
    const lastCol = employeeSheet.getLastColumn();

    // Get default access values for this employee classification
    const defaultAccessValues = getDefaultAccessValues(defaultAccessSheet, jobClassification);
    
    if (!defaultAccessValues) {
        Logger.log(`Warning: No default access values found for classification "${jobClassification}". Using empty values.`);
    }

    // --- Perform Copy/Paste and Update ---
    const newRowIndex = lastRow + 1;

    // Copy formatting from the row above if it exists
    if (lastRow >= 2) { // Ensure there's a data row to copy format from
        const sourceRange = employeeSheet.getRange(lastRow, 1, 1, lastCol);
        const targetRange = employeeSheet.getRange(newRowIndex, 1, 1, lastCol);
        
        // Copy both format and data validation (for checkboxes)
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
        
        // Clear content from the new row AFTER pasting format and validation
        targetRange.clearContent();
    }

    // Column mapping based on the Employee sheet structure (1-based indexing)
    // Note: Column A is a buffer column, actual data starts at column B
    const columnMapping = {
        nameOfIndividual: 2,           // B - Name of Individual
        employeeClassification: 3,      // C - Employee Classification 
        noAccess: 4,                   // D - No Access
        healthFinancialInfo: 5,        // E - Health & financial information (EMR & PMS)
        emailCloudDrive: 6,            // F - Email and Cloud Drive
        textVoicemail: 7,              // G - Text and Voicemail (Klara)
        voip: 8,                       // H - VOIP
        rems: 9,                       // I - REMS (iPledge, etc.)
        officeKey: 10,                 // J - Office Key
        alarmSystemCode: 11,           // K - Alarm System Code
        dateAccessGranted: 12,         // L - Date Access Granted
        dateAccessRevoked: 13,         // M - Date Access Revoked
        comments: 14                   // N - Comments
    };

    // Debug logging
    Logger.log(`Setting values for row ${newRowIndex}:`);
    Logger.log(`  Employee Name: "${employeeName}" -> Column ${columnMapping.nameOfIndividual} (B)`);
    Logger.log(`  Job Classification: "${jobClassification}" -> Column ${columnMapping.employeeClassification} (C)`);
    Logger.log(`  Start Date: "${startDate}" -> Column ${columnMapping.dateAccessGranted} (L)`);
    
    // Set the employee name and classification
    employeeSheet.getRange(newRowIndex, columnMapping.nameOfIndividual).setValue(employeeName);
    employeeSheet.getRange(newRowIndex, columnMapping.employeeClassification).setValue(jobClassification);
    
    // Set access values from default access lookup (copy values, not formulas)
    if (defaultAccessValues) {
        // Set checkbox values explicitly with proper boolean conversion
        const accessColumns = [
            { col: columnMapping.noAccess, value: defaultAccessValues.noAccess },
            { col: columnMapping.healthFinancialInfo, value: defaultAccessValues.healthFinancialInfo },
            { col: columnMapping.emailCloudDrive, value: defaultAccessValues.emailCloudDrive },
            { col: columnMapping.textVoicemail, value: defaultAccessValues.textVoicemail },
            { col: columnMapping.voip, value: defaultAccessValues.voip },
            { col: columnMapping.rems, value: defaultAccessValues.rems }
        ];
        
        accessColumns.forEach(item => {
            const cellRange = employeeSheet.getRange(newRowIndex, item.col);
            // Ensure checkbox data validation is set
            const validation = SpreadsheetApp.newDataValidation()
                .requireCheckbox()
                .build();
            cellRange.setDataValidation(validation);
            // Set the boolean value
            cellRange.setValue(item.value === true);
        });
        
        Logger.log(`  Default access values applied: ${JSON.stringify(defaultAccessValues)}`);
    }
    
    // Set Office Key and Alarm System Code to false (unchecked) with checkbox validation
    const officeKeyRange = employeeSheet.getRange(newRowIndex, columnMapping.officeKey);
    const alarmCodeRange = employeeSheet.getRange(newRowIndex, columnMapping.alarmSystemCode);
    
    // Ensure checkbox data validation
    const checkboxValidation = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .build();
    
    officeKeyRange.setDataValidation(checkboxValidation);
    alarmCodeRange.setDataValidation(checkboxValidation);
    
    officeKeyRange.setValue(false);
    alarmCodeRange.setValue(false);
    
    // Set the date access granted to the onboarding start date
    employeeSheet.getRange(newRowIndex, columnMapping.dateAccessGranted).setValue(startDate);

    Logger.log(`Updated Access Log sheet "${sheetName}" for ${employeeName} with classification "${jobClassification}".`);

  } catch (error) {
    Logger.log(`Error updating Access Log sheet "${sheetName}" for ${employeeName}: ${error.message} \nStack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`Failed to update the Access Log for ${employeeName}. Check logs.`);
  }
}