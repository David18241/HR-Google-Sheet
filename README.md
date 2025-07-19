# HR Management Google Apps Script

A comprehensive Google Apps Script solution for automating HR workflows including employee onboarding, offboarding, training attestations, and document management integrated with Google Sheets, Drive, and Gmail.

## 🚀 Features

- **Employee Onboarding**: Automated folder creation, document generation, and group assignments
- **Employee Offboarding**: Secure access removal, folder archiving, and group cleanup  
- **Training Management**: OSHA and HIPAA training attestation workflows
- **Access Log Management**: Automated HIPAA access logging with role-based permissions
- **Document Workflows**: Template-based document generation with personalization
- **Email Automation**: Automated welcome emails, notifications, and training reminders

## 📋 Prerequisites

- Google Workspace account with admin privileges
- Google Apps Script project
- Required Google APIs enabled:
  - Admin Directory API ✅
  - Drive API v3
  - Docs API v1
  - Gmail API v1

## 🛠️ Installation

### 1. Clone and Setup

```bash
git clone https://github.com/David18241/HR-Google-Sheet.git
cd HR-Google-Sheet
```

### 2. Configure Your Environment

1. **Copy configuration templates:**
   ```bash
   cp Constants.example.js Constants.js
   cp .clasp.example.json .clasp.json
   ```

2. **Update Constants.js with your values:**
   - Replace all `YOUR_*` placeholders with actual Google Drive IDs
   - Update email addresses to match your organization
   - Customize spreadsheet column headers if needed

3. **Update .clasp.json:**
   - Replace `YOUR_GOOGLE_APPS_SCRIPT_PROJECT_ID` with your Apps Script project ID

### 3. Deploy to Google Apps Script

If using [clasp](https://github.com/google/clasp):

```bash
npm install -g @google/clasp
clasp login
clasp push
```

Or manually copy the .js files to your Google Apps Script project.

### 4. Enable Required Services

In your Google Apps Script project:
1. Go to Services in the left sidebar
2. Add and enable:
   - Admin Directory API
   - Drive API
   - Docs API  
   - Gmail API

## 📊 Spreadsheet Setup

### Main Personnel Sheet
Required columns:
- First Name
- Last Name  
- Work Email
- Personal Email
- Start Date
- End Date
- Primary Classification
- Group Email
- Active
- Job Folder ID
- Employee Folder ID
- Employee MedRec Folder ID

### Access Log Spreadsheet
Required sheets:
- **Employees**: Main access tracking
- **Default Access**: Role-based access templates

## 🔧 Usage

### Onboarding Workflow
1. Add new employee to Personnel sheet
2. Select the employee row
3. Run: `Utilities > Onboard Employee`

### Offboarding Workflow  
1. Select employee row in Personnel sheet
2. Run: `Utilities > Offboard Employee`
3. Confirm the action in the dialog

### Access Log Management
The system automatically updates the HIPAA access log during onboarding based on employee classification and default access permissions.

## 🛡️ Security Features

- **HIPAA Compliance**: Automated access logging and audit trails
- **Role-based Access**: Default permissions based on job classification
- **Secure Offboarding**: Comprehensive access removal and folder archiving
- **Error Handling**: Robust error handling with retry logic for API calls
- **Audit Logging**: Detailed logging for all operations

## 🏗️ Architecture

### Core Components

- **Constants.js**: Central configuration
- **Utilities.js**: Main workflow orchestration  
- **WorkflowHelpers.js**: Modular workflow functions
- **AdminHelpers.js**: Google Admin Directory integration
- **DriveHelpers.js**: Google Drive operations
- **EmailHelpers.js**: Email generation and sending
- **SheetHelpers.js**: Google Sheets utilities
- **ErrorHandling.js**: Centralized error handling

### Key Features

- **Modular Design**: Separated concerns for easy maintenance
- **Error Recovery**: Automatic retry logic with exponential backoff
- **Progress Tracking**: User-friendly progress indicators
- **Comprehensive Logging**: Detailed operation logs for debugging

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📝 License

This project is provided as-is for educational and organizational use. Please ensure compliance with your organization's security and privacy policies.

## ⚠️ Important Notes

- **Test First**: Always test in a non-production environment
- **Backup Data**: Ensure you have backups before running workflows
- **Review Permissions**: Verify Google API permissions meet your security requirements
- **Customize**: Adapt the code to match your organization's specific requirements

## 🆘 Support

For issues and questions:
1. Check the logs in Google Apps Script editor
2. Review the error handling documentation
3. Submit an issue on GitHub

---

🤖 *This project includes components generated with [Claude Code](https://claude.ai/code)*