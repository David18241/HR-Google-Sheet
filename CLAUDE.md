# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script project for HR management automation at a healthcare practice. The system manages employee onboarding, offboarding, training attestations, and document workflows through Google Sheets, Drive, and Gmail integration.

## Architecture

### Core Components

- **Constants.js** - Central configuration file containing:
  - Spreadsheet and folder IDs
  - Google Doc template IDs for forms and checklists
  - Email template IDs
  - Email addresses and group configurations
  - Column header mappings for the Personnel sheet

- **SheetHelpers.js** - Google Sheets utilities:
  - `getSheetDataWithHeaders()` - Gets sheet data with header mappings
  - `getActiveRowData()` - Validates and extracts currently selected row data
  - `parseEmployeeName()` - Parses "Last, First" name format
  - `formatDate()` - Formats dates for display

- **DriveHelpers.js** - Google Drive file/folder operations:
  - `createFolderIfNotExists()` - Creates folders with duplicate checking
  - `copyAndPersonalizeDocument()` - Clones templates with placeholder replacement
  - `replacePlaceholderWithLink()` - Replaces text with hyperlinks in documents
  - Permission management functions for files and folders

- **EmailHelpers.js** - Email generation and sending:
  - `createEmailFromTemplate()` - Main email workflow function
  - `convertBodyToHtml()` - Converts Google Docs to HTML for email
  - Comprehensive Google Doc to HTML conversion functions

- **AdminHelpers.js** - Google Admin Directory integration:
  - `addMembershipToGroups()` - Adds users to Google Groups
  - `removeMembershipFromGroups()` - Removes users from Google Groups  
  - `updateAccessLog()` - Updates HIPAA access tracking spreadsheet

- **Utilities.js** - Main workflow orchestration:
  - `onboardEmployeeWorkflow()` - Complete new employee setup
  - `offboardEmployeeWorkflow()` - Employee termination process
  - Training attestation email workflows for OSHA/HIPAA compliance
  - `createBeforeFirstDayEmailWorkflow()` - Pre-start communication

### Data Flow

1. Employee data is managed in the main Personnel sheet
2. Workflows are triggered by selecting a row and using the custom menu
3. The system creates personalized folders, documents, and emails
4. Updates are made to tracking spreadsheets and group memberships
5. Notifications are sent to relevant parties

## Common Development Tasks

### Testing Workflows
- Use Google Apps Script editor's built-in debugger
- Check execution logs via `View > Logs` in the Apps Script editor
- Test with non-production spreadsheet data when possible

### Deployment
```bash
# If using clasp (Google Apps Script CLI)
clasp push
```

### Key Dependencies
- Requires enabled Google Apps Script Advanced Services:
  - Admin Directory API (for group management)
  - Drive API v3
  - Docs API v1
  - Gmail API v1
- Must be run by users with appropriate Google Workspace admin permissions

### Security Considerations
- All template IDs and folder IDs are stored in Constants.js
- Email addresses and group configurations are centrally managed
- HIPAA access logging is automatically maintained
- Document sharing permissions are programmatically controlled

### Error Handling
- All major functions include try-catch blocks with detailed logging
- User-facing alerts for common error scenarios
- Graceful degradation when optional features fail
- Cleanup of temporary files on errors