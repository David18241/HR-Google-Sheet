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

     // Remove from specific group if provided
    if (specificGroupEmail) {
        if (!checkServiceAvailability("AdminDirectory", AdminDirectory)) {
            return;
        }
        
        const removeFromGroupOperation = () => {
            AdminDirectory.Members.remove(specificGroupEmail, userEmail);
            return true;
        };
        
        const result = executeWithRetry(
            removeFromGroupOperation, 
            `Removing ${userEmail} from group ${specificGroupEmail}`
        );
        
        if (result) {
            Logger.log(`User ${userEmail} removed from the group ${specificGroupEmail}.`);
        } else {
            Logger.log(`Failed to remove ${userEmail} from group ${specificGroupEmail} after retries.`);
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
 * Updates the Access Log spreadsheet with new employee information.
 * Copies formatting from the last row and clears specific columns before adding new data.
 *
 * @param {string} spreadsheetId The ID of the Access Log spreadsheet.
 * @param {string} sheetName The name of the sheet within the Access Log (e.g., "Employees").
 * @param {string} employeeName Formatted employee name ("Last, First").
 * @param {string} startDate Formatted start date string. // Changed type hint to string as passed from main workflow
 * @param {string} jobClassification The employee's job title/classification.
 */
function updateAccessLog(spreadsheetId, sheetName, employeeName, startDate, jobClassification) {
  try {
    const accessLogSs = SpreadsheetApp.openById(spreadsheetId);
    const accessLogSheet = accessLogSs.getSheetByName(sheetName);

    if (!accessLogSheet) {
        Logger.log(`Error: Access Log sheet "${sheetName}" not found in spreadsheet ID ${spreadsheetId}.`);
        SpreadsheetApp.getUi().alert(`Error: Could not find the sheet "${sheetName}" in the Access Log. Update skipped.`);
        return;
    }

    const lastRow = accessLogSheet.getLastRow();
    const lastCol = accessLogSheet.getLastColumn();

    // Ensure there's data to copy from
    if (lastRow < 2) { // Assuming header is row 1
        Logger.log("Access Log sheet has no data rows to copy format from. Appending data without format copy.");
         // Adjust column order/number based on your ACTUAL Access Log sheet structure
         // Example: Assuming Name=Col B, Job=Col C, StartDate=Col F
         const newRowData = [];
         newRowData[1] = employeeName; // Index 1 for Col B
         newRowData[2] = jobClassification; // Index 2 for Col C
         newRowData[5] = startDate; // Index 5 for Col F (adjust!)
         accessLogSheet.appendRow(newRowData);
         Logger.log(`Appended new row to Access Log for ${employeeName} as sheet was empty/had only header.`);
         return;
    }

    // --- Define columns to update ---
    // It's safer to find columns by header dynamically if possible.
    // Using fixed indices from original script for now, but verify these match YOUR Access Log sheet:
    const employeeNameCol = 2; // B
    const jobClassCol = 3;     // C
    // Original script set date in lastCol - 2. Let's assume this is correct, BUT VERIFY.
    // Example: If lastCol=8 (H), this is F (col 6)
    const startDateCol = lastCol - 2;

    // --- Perform Copy/Paste and Update ---
    const newRowIndex = lastRow + 1;

    // Copy the entire previous data row's format to the new row
    const sourceRange = accessLogSheet.getRange(lastRow, 1, 1, lastCol);
    const targetRange = accessLogSheet.getRange(newRowIndex, 1, 1, lastCol);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    // Clear content from the new row AFTER pasting format
    targetRange.clearContent();

    // Set new values in the specific columns for the new row
    accessLogSheet.getRange(newRowIndex, employeeNameCol).setValue(employeeName);
    accessLogSheet.getRange(newRowIndex, jobClassCol).setValue(jobClassification);
    // Ensure startDate is a value Sheets can understand (string "Month Day, Year" or a JS Date object)
    accessLogSheet.getRange(newRowIndex, startDateCol).setValue(startDate);

    // Original script cleared last 5 columns - let's replicate ONLY if necessary.
    // This seems odd if you just set the date in lastCol - 2.
    // If you truly need to clear specific trailing columns AFTER setting values, do it here.
    // Example: Clear cols lastCol-4 through lastCol
    // const startClearCol = Math.max(1, lastCol - 4); // Don't try to clear before Col A
    // if (lastCol >= startClearCol) {
    //    accessLogSheet.getRange(newRowIndex, startClearCol, 1, lastCol - startClearCol + 1).clearContent();
    // }

    Logger.log(`Updated Access Log sheet "${sheetName}" for ${employeeName}.`);

  } catch (error) {
    Logger.log(`Error updating Access Log sheet "${sheetName}" for ${employeeName}: ${error.message} \nStack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`Failed to update the Access Log for ${employeeName}. Check logs.`);
  }
}