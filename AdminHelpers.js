/**
 * @fileoverview Helper functions using the Admin SDK Directory Service.
 *               Requires 'AdminDirectory' advanced service to be enabled.
 */

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

  // Add to specific group if provided
  if (specificGroupEmail) {
    try {
      // Check if AdminDirectory service is available
      if (typeof AdminDirectory === 'undefined') {
          throw new Error("AdminDirectory service is not enabled or available.");
      }
      AdminDirectory.Members.insert(member, specificGroupEmail);
      Logger.log(`User ${userEmail} added as a member to the group ${specificGroupEmail}.`);
    } catch (err) {
      // Check for specific error messages
      if (err.message.includes("Member already exists")) {
        Logger.log(`User ${userEmail} is already in the group ${specificGroupEmail}.`);
      } else if (err.message.includes("Resource Not Found: groupKey")) {
         Logger.log(`Error adding ${userEmail} to ${specificGroupEmail}: Group not found.`);
         SpreadsheetApp.getUi().alert(`Could not add user to group "${specificGroupEmail}" (Group not found).`);
      } else if (err.message.includes("Not Authorized to access this resource/api")) {
         Logger.log(`Authorization error adding ${userEmail} to ${specificGroupEmail}: ${err.message}`);
         SpreadsheetApp.getUi().alert(`Authorization error adding user to group "${specificGroupEmail}". Ensure script has necessary permissions.`);
      }
      else {
        Logger.log(`Error adding ${userEmail} to group ${specificGroupEmail}: ${err.message}`);
        // Consider alerting the UI for unexpected errors
        SpreadsheetApp.getUi().alert(`An unexpected error occurred adding the user to group "${specificGroupEmail}". Check logs.`);
      }
    }
  } else {
      Logger.log(`Skipping adding ${userEmail} to specific group (no group email provided).`);
  }

  // Add to practice-wide group
  if (practiceWideGroupEmail) {
     try {
       if (typeof AdminDirectory === 'undefined') {
           throw new Error("AdminDirectory service is not enabled or available.");
       }
      AdminDirectory.Members.insert(member, practiceWideGroupEmail);
      Logger.log(`User ${userEmail} added as a member to the group ${practiceWideGroupEmail}.`);
    } catch (err) {
      if (err.message.includes("Member already exists")) {
        Logger.log(`User ${userEmail} is already in the group ${practiceWideGroupEmail}.`);
      } else if (err.message.includes("Resource Not Found: groupKey")) {
         Logger.log(`Error adding ${userEmail} to ${practiceWideGroupEmail}: Group not found.`);
          SpreadsheetApp.getUi().alert(`Could not add user to practice group "${practiceWideGroupEmail}" (Group not found).`);
      } else if (err.message.includes("Not Authorized to access this resource/api")) {
         Logger.log(`Authorization error adding ${userEmail} to ${practiceWideGroupEmail}: ${err.message}`);
         SpreadsheetApp.getUi().alert(`Authorization error adding user to practice group "${practiceWideGroupEmail}". Ensure script has necessary permissions.`);
      } else {
        Logger.log(`Error adding ${userEmail} to practice group ${practiceWideGroupEmail}: ${err.message}`);
         SpreadsheetApp.getUi().alert(`An unexpected error occurred adding the user to practice group "${practiceWideGroupEmail}". Check logs.`);
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
        try {
             if (typeof AdminDirectory === 'undefined') {
                throw new Error("AdminDirectory service is not enabled or available.");
            }
            AdminDirectory.Members.remove(specificGroupEmail, userEmail);
            Logger.log(`User ${userEmail} removed from the group ${specificGroupEmail}.`);
        } catch (err) {
            if (err.message.includes("Resource Not Found: memberKey")) {
                Logger.log(`User ${userEmail} was not found in the group ${specificGroupEmail}.`);
            } else if (err.message.includes("Resource Not Found: groupKey")) {
                 Logger.log(`Error removing ${userEmail} from ${specificGroupEmail}: Group not found.`);
                 SpreadsheetApp.getUi().alert(`Could not remove user from group "${specificGroupEmail}" (Group not found).`);
            } else if (err.message.includes("Not Authorized to access this resource/api")) {
                Logger.log(`Authorization error removing ${userEmail} from ${specificGroupEmail}: ${err.message}`);
                SpreadsheetApp.getUi().alert(`Authorization error removing user from group "${specificGroupEmail}". Ensure script has necessary permissions.`);
           } else {
                Logger.log(`Error removing ${userEmail} from group ${specificGroupEmail}: ${err.message}`);
                SpreadsheetApp.getUi().alert(`An unexpected error occurred removing the user from group "${specificGroupEmail}". Check logs.`);
            }
        }
    } else {
      Logger.log(`Skipping removing ${userEmail} from specific group (no group email provided).`);
    }

    // Remove from practice-wide group
    if (practiceWideGroupEmail) {
       try {
            if (typeof AdminDirectory === 'undefined') {
                throw new Error("AdminDirectory service is not enabled or available.");
            }
            AdminDirectory.Members.remove(practiceWideGroupEmail, userEmail);
            Logger.log(`User ${userEmail} removed from the group ${practiceWideGroupEmail}.`);
        } catch (err) {
            if (err.message.includes("Resource Not Found: memberKey")) {
                Logger.log(`User ${userEmail} was not found in the group ${practiceWideGroupEmail}.`);
            } else if (err.message.includes("Resource Not Found: groupKey")) {
                 Logger.log(`Error removing ${userEmail} from ${practiceWideGroupEmail}: Group not found.`);
                 SpreadsheetApp.getUi().alert(`Could not remove user from practice group "${practiceWideGroupEmail}" (Group not found).`);
            } else if (err.message.includes("Not Authorized to access this resource/api")) {
                Logger.log(`Authorization error removing ${userEmail} from ${practiceWideGroupEmail}: ${err.message}`);
                SpreadsheetApp.getUi().alert(`Authorization error removing user from practice group "${practiceWideGroupEmail}". Ensure script has necessary permissions.`);
            } else {
                Logger.log(`Error removing ${userEmail} from practice group ${practiceWideGroupEmail}: ${err.message}`);
                SpreadsheetApp.getUi().alert(`An unexpected error occurred removing the user from practice group "${practiceWideGroupEmail}". Check logs.`);
            }
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