/**
 * @fileoverview Centralized error handling and utility functions.
 */

/**
 * Standardized error handling for workflow functions.
 * Logs the error and shows user-friendly alert.
 *
 * @param {Error} error The error object
 * @param {string} context Context description (e.g., "Onboarding workflow")
 * @param {string} userMessage User-friendly message to display
 * @param {boolean} [showAlert=true] Whether to show UI alert
 * @returns {false} Always returns false to indicate failure
 */
function handleWorkflowError(error, context, userMessage, showAlert = true) {
  const errorMessage = `${context}: ${error.message}`;
  const stackTrace = error.stack || 'No stack trace available';
  
  Logger.log(`ERROR - ${errorMessage}\nStack: ${stackTrace}`);
  
  if (showAlert) {
    SpreadsheetApp.getUi().alert(`${userMessage}\n\nCheck logs for technical details.`);
  }
  
  return false;
}

/**
 * Handles Google API service availability check.
 * Returns true if service is available, false otherwise.
 *
 * @param {string} serviceName Name of the service (e.g., "AdminDirectory")
 * @param {Object} serviceObject The service object to check
 * @returns {boolean} True if service is available
 */
function checkServiceAvailability(serviceName, serviceObject) {
  if (typeof serviceObject === 'undefined') {
    const message = `${serviceName} service is not enabled or available. Please enable it in the Apps Script project settings.`;
    Logger.log(`SERVICE ERROR: ${message}`);
    SpreadsheetApp.getUi().alert(`Service Error: ${message}`);
    return false;
  }
  return true;
}

/**
 * Wraps Google API calls with retry logic and error handling.
 * 
 * @param {Function} apiCall Function that makes the API call
 * @param {string} operationName Description of the operation for logging
 * @param {number} [maxRetries=3] Maximum number of retry attempts
 * @param {number} [initialDelay=1000] Initial delay in milliseconds
 * @returns {*} Result of the API call or null if all retries failed
 */
function executeWithRetry(apiCall, operationName, maxRetries = 3, initialDelay = 1000) {
  let lastError;
  let delay = initialDelay;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      Logger.log(`Attempt ${attempt}/${maxRetries} for: ${operationName}`);
      return apiCall();
    } catch (error) {
      lastError = error;
      Logger.log(`Attempt ${attempt} failed for ${operationName}: ${error.message}`);
      
      // Check if it's a quota/rate limit error that might benefit from retry
      if (error.message.includes('rate limit') || error.message.includes('quota') || 
          error.message.includes('timeout') || error.message.includes('service unavailable')) {
        
        if (attempt < maxRetries) {
          Logger.log(`Retrying ${operationName} in ${delay}ms...`);
          Utilities.sleep(delay);
          delay *= 2; // Exponential backoff
        }
      } else {
        // Non-retryable error, fail immediately
        Logger.log(`Non-retryable error for ${operationName}, failing immediately`);
        break;
      }
    }
  }
  
  Logger.log(`All retry attempts failed for ${operationName}: ${lastError.message}`);
  return null;
}

/**
 * Validates that required fields are present and not empty.
 *
 * @param {Object} data Object containing data to validate
 * @param {Array<string>} requiredFields Array of required field names
 * @param {string} [context="Data validation"] Context for error messages
 * @returns {{isValid: boolean, missingFields: Array<string>}} Validation result
 */
function validateRequiredFields(data, requiredFields, context = "Data validation") {
  const missingFields = [];
  
  for (const field of requiredFields) {
    if (!data[field] || (typeof data[field] === 'string' && data[field].trim() === '')) {
      missingFields.push(field);
    }
  }
  
  const isValid = missingFields.length === 0;
  
  if (!isValid) {
    Logger.log(`${context} failed: Missing required fields: ${missingFields.join(', ')}`);
  }
  
  return { isValid, missingFields };
}

/**
 * Safely attempts to clean up a temporary file.
 *
 * @param {string} fileId Google Drive file ID to clean up
 * @param {string} [context="File cleanup"] Context for logging
 */
function safeFileCleanup(fileId, context = "File cleanup") {
  if (!fileId) return;
  
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    Logger.log(`${context}: Successfully cleaned up file ${fileId}`);
  } catch (cleanupError) {
    Logger.log(`${context}: Failed to clean up file ${fileId}: ${cleanupError.message}`);
    // Don't throw - cleanup failures shouldn't break main workflow
  }
}

/**
 * Shows a progress sidebar to the user.
 *
 * @param {string} title Title for the sidebar
 * @param {string} message Current status message
 * @param {number} [current] Current progress number
 * @param {number} [total] Total number of items
 */
function showProgressSidebar(title, message, current = null, total = null) {
  let content = `<p>${message}</p>`;
  
  if (current !== null && total !== null) {
    const percentage = Math.round((current / total) * 100);
    content += `<p>Progress: ${current}/${total} (${percentage}%)</p>`;
    content += `<div style="width: 100%; background-color: #f0f0f0; border-radius: 3px;">
                  <div style="width: ${percentage}%; background-color: #4CAF50; height: 20px; border-radius: 3px;"></div>
                </div>`;
  }
  
  const htmlOutput = HtmlService.createHtmlOutput(content)
    .setWidth(300)
    .setHeight(150);
  
  SpreadsheetApp.getUi().showSidebar(htmlOutput.setTitle(title));
  SpreadsheetApp.flush();
}

/**
 * Closes any open sidebar.
 */
function closeSidebar() {
  try {
    // There's no direct method to close sidebar, but we can replace it with empty content
    const htmlOutput = HtmlService.createHtmlOutput('')
      .setWidth(1)
      .setHeight(1);
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  } catch (error) {
    // Ignore errors when closing sidebar
    Logger.log(`Failed to close sidebar: ${error.message}`);
  }
}

/**
 * Validates email address format.
 *
 * @param {string} email Email address to validate
 * @returns {boolean} True if email format is valid
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email.trim());
}

/**
 * Validates date object and ensures it's not in the past (for start dates).
 *
 * @param {Date} dateObj Date object to validate
 * @param {boolean} [allowPastDates=true] Whether to allow past dates
 * @returns {{isValid: boolean, message: string}} Validation result
 */
function validateDate(dateObj, allowPastDates = true) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
    return { isValid: false, message: "Invalid date object" };
  }
  
  if (!allowPastDates && dateObj < new Date()) {
    return { isValid: false, message: "Date cannot be in the past" };
  }
  
  return { isValid: true, message: "Valid date" };
}

/**
 * Creates a standardized confirmation dialog.
 *
 * @param {string} title Dialog title
 * @param {string} message Dialog message
 * @param {string} confirmationText Text user must type to confirm
 * @returns {boolean} True if user confirmed correctly
 */
function showConfirmationDialog(title, message, confirmationText) {
  const ui = SpreadsheetApp.getUi();
  const fullMessage = `${message}\n\nType "${confirmationText}" to confirm:`;
  
  const response = ui.prompt(title, fullMessage, ui.ButtonSet.OK_CANCEL);
  
  return response.getSelectedButton() === ui.Button.OK && 
         response.getResponseText().trim().toUpperCase() === confirmationText.toUpperCase();
}