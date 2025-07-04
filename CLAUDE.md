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
  - Admin Directory API (for group management) ✅ **Now enabled**
  - Drive API v3
  - Docs API v1
  - Gmail API v1
- Must be run by users with appropriate Google Workspace admin permissions

### Recent Improvements (2024)

#### **Standardized Error Handling** (ErrorHandling.js)
- `handleWorkflowError()` - Centralized error logging and user notifications
- `executeWithRetry()` - Automatic retry logic for Google API calls with exponential backoff
- `validateRequiredFields()` - Input validation with detailed error reporting
- `showProgressSidebar()` - User-friendly progress indicators for long operations

#### **Modular Workflow Architecture** (WorkflowHelpers.js)
- Large workflow functions broken into smaller, testable units
- `extractAndValidateEmployeeData()` - Centralized data extraction and validation
- Separate functions for folder creation, document generation, permissions, etc.
- Improved error recovery and progress tracking

#### **Complete Offboarding Implementation**
- Fully implemented offboarding workflow with confirmation dialogs
- Automatic group removal and folder access management
- Optional offboarding documentation and email generation
- Progress tracking and comprehensive error handling

#### **Enhanced API Reliability**
- Retry logic for Google API calls handles rate limits and temporary failures
- Service availability checks before API operations
- Graceful degradation when services are unavailable

### Security Considerations
- All template IDs and folder IDs are stored in Constants.js
- Email addresses and group configurations are centrally managed
- HIPAA access logging is automatically maintained
- Document sharing permissions are programmatically controlled
- Confirmation dialogs for destructive operations (offboarding)

### Error Handling
- Centralized error handling with consistent patterns across all modules
- User-facing progress indicators and status updates
- Detailed logging with context for debugging
- Automatic cleanup of temporary files on errors
- Retry logic for transient API failures