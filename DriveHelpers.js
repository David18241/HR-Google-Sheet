/**
 * @fileoverview Helper functions for Google Drive interactions (folders, documents).
 */

/**
 * Creates a folder within a parent folder if it doesn't already exist.
 *
 * @param {string} folderName The name of the folder to create.
 * @param {string} parentFolderId The ID of the parent folder.
 * @returns {Folder|null} The created or existing Folder object, or null on error.
 */
function createFolderIfNotExists(folderName, parentFolderId) {
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const existingFolders = parentFolder.getFoldersByName(folderName);

    if (existingFolders.hasNext()) {
      const existingFolder = existingFolders.next();
      Logger.log(`Folder "${folderName}" already exists in parent ID ${parentFolderId}: ${existingFolder.getUrl()}`);
      return existingFolder;
    } else {
      const newFolder = parentFolder.createFolder(folderName);
      Logger.log(`Created folder "${folderName}" in parent ID ${parentFolderId}: ${newFolder.getUrl()}`);
      return newFolder;
    }
  } catch (error) {
    Logger.log(`Error creating/finding folder "${folderName}" in parent ID ${parentFolderId}: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error creating folder "${folderName}". Check logs for details.`);
    return null;
  }
}

/**
 * Copies a template document, replaces placeholders, and saves it in a destination folder.
 *
 * @param {string} templateId The ID of the Google Doc template.
 * @param {Folder} destinationFolder The Drive Folder object to save the new document in.
 * @param {string} newFileName The desired name for the new document.
 * @param {Object} placeholderValues An object where keys are placeholders (e.g., "{{FirstName}}") and values are the replacements.
 * @returns {Document|null} The newly created Google Document object, or null on error/template not found.
 */
function copyAndPersonalizeDocument(templateId, destinationFolder, newFileName, placeholderValues) {
  try {
    const templateFile = DriveApp.getFileById(templateId);
    const newFile = templateFile.makeCopy(newFileName, destinationFolder);
    const newDoc = DocumentApp.openById(newFile.getId());
    const body = newDoc.getBody();

    for (const placeholder in placeholderValues) {
      if (placeholderValues.hasOwnProperty(placeholder)) {
        // Simple text replacement
        body.replaceText(placeholder, placeholderValues[placeholder] || ''); // Replace with empty string if value is null/undefined
      }
    }
    newDoc.saveAndClose();
    Logger.log(`Created and personalized document "${newFileName}" from template ID ${templateId}. URL: ${newFile.getUrl()}`);
    return newDoc; // Return the Document object for potential further manipulation (like adding editors)

  } catch (e) {
      // Handle specific error: If template ID is invalid
      if (e.message.includes("Not Found") || e.message.includes("Unable to retrieve file")) {
          Logger.log(`Error: Template document with ID "${templateId}" not found or inaccessible.`);
          SpreadsheetApp.getUi().alert(`Error: Could not find the template document needed for "${newFileName}". Please check the Template ID (${templateId}).`);
      } else {
          Logger.log(`Error copying or personalizing document "${newFileName}" from template ID ${templateId}: ${e.message} \nStack: ${e.stack}`);
          SpreadsheetApp.getUi().alert(`Error creating document "${newFileName}". Check logs for details.`);
      }
      // Attempt to trash the potentially partially created file if it exists
      try {
          const tempFile = DriveApp.getFilesByName(newFileName).next(); // Simple check, might not be robust if names collide
          if (tempFile) tempFile.setTrashed(true);
      } catch (cleanupError) { /* Ignore cleanup errors */ }

      return null;
  }
}


/**
 * Finds a text placeholder in a Google Doc body and replaces it with hyperlinked text.
 * Handles cases where the placeholder might appear multiple times.
 *
 * @param {Body} body The Google Doc Body object.
 * @param {string} placeholder The text to find (e.g., "{{myLink}}").
 * @param {string} linkText The text to display for the link.
 * @param {string} url The URL for the hyperlink.
 */
function replacePlaceholderWithLink(body, placeholder, linkText, url) {
  let foundElement = body.findText(placeholder);
  while (foundElement) {
    const element = foundElement.getElement();
    if (element.getType() === DocumentApp.ElementType.TEXT) {
      const textElement = element.asText();
      const startOffset = foundElement.getStartOffset();
      const endOffsetInclusive = foundElement.getEndOffsetInclusive();

      // Replace the text first
      textElement.deleteText(startOffset, endOffsetInclusive);
      textElement.insertText(startOffset, linkText);

      // Apply the link to the newly inserted text
      // Important: Use the new offsets after insertion
      textElement.setLinkUrl(startOffset, startOffset + linkText.length - 1, url);

      // Find the next occurrence, starting search after the current one
      // Use findText(pattern, from) - 'from' is crucial here
      foundElement = body.findText(placeholder, foundElement);
    } else {
      // Placeholder found but not in a simple Text element, skip or log warning
      Logger.log(`Warning: Placeholder "${placeholder}" found within a non-text element type (${element.getType()}). Skipping link replacement.`);
      // To prevent infinite loop if findText keeps finding the same non-text element:
      foundElement = body.findText(placeholder, foundElement);
    }
  }
}

/**
 * Adds a user as an editor to a Google Drive file.
 * Includes basic error handling.
 *
 * @param {File|Document} fileObject The File or Document object.
 * @param {string} email The email address of the user to add.
 */
function addEditorToFile(fileObject, email) {
    if (!fileObject || !email) {
        Logger.log(`Skipping adding editor: Invalid file object or email (${email}).`);
        return;
    }
    try {
        fileObject.addEditor(email);
        Logger.log(`Added ${email} as editor to file ID ${fileObject.getId()}.`);
    } catch (error) {
        Logger.log(`Error adding ${email} as editor to file ID ${fileObject.getId()}: ${error.message}`);
        // Optionally alert the user, but might be too noisy
        // SpreadsheetApp.getUi().alert(`Could not share file with ${email}. Please check permissions or share manually.`);
    }
}

/**
 * Adds a user as a viewer to a Google Drive folder.
 * Includes basic error handling.
 *
 * @param {Folder} folderObject The Folder object.
 * @param {string} email The email address of the user to add.
 */
function addViewerToFolder(folderObject, email) {
    if (!folderObject || !email) {
        Logger.log(`Skipping adding viewer: Invalid folder object or email (${email}).`);
        return;
    }
    try {
        folderObject.addViewer(email);
        Logger.log(`Added ${email} as viewer to folder ID ${folderObject.getId()}.`);
    } catch (error) {
        Logger.log(`Error adding ${email} as viewer to folder ID ${folderObject.getId()}: ${error.message}`);
        // Optionally alert the user
        // SpreadsheetApp.getUi().alert(`Could not share folder with ${email}. Please check permissions or share manually.`);
    }
}

/**
 * Adds a user as a commenter to a Google Drive folder.
 * Includes basic error handling.
 *
 * @param {Folder} folderObject The Folder object.
 * @param {string} email The email address of the user to add.
 */
function addCommenterToFolder(folderObject, email) {
    if (!folderObject || !email) {
        Logger.log(`Skipping adding commenter: Invalid folder object or email (${email}).`);
        return;
    }
    try {
        const folderId = folderObject.getId();
        
        // First try using the Advanced Drive Service for true "commenter" access
        try {
            const permission = {
                'role': 'commenter',
                'type': 'user',
                'emailAddress': email
            };
            
            Drive.Permissions.create(permission, folderId, {
                'sendNotificationEmails': false,
                'supportsAllDrives': true  // Required for shared drives
            });
            Logger.log(`Added ${email} as commenter to folder ID ${folderId} using Drive API.`);
        } catch (driveApiError) {
            // If Drive API fails, fall back to addViewer
            Logger.log(`Drive API failed, falling back to addViewer: ${driveApiError.message}`);
            folderObject.addViewer(email);
            Logger.log(`Added ${email} as viewer to folder ID ${folderId} (fallback).`);
        }
    } catch (error) {
        Logger.log(`Error adding ${email} as commenter to folder ID ${folderObject.getId()}: ${error.message}`);
        // Optionally alert the user
        // SpreadsheetApp.getUi().alert(`Could not share folder with ${email}. Please check permissions or share manually.`);
    }
}