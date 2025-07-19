/**
 * @fileoverview Example configuration file for HR Management Script.
 * Copy this to Constants.js and replace placeholders with your actual values.
 */

// --- Spreadsheet / Folder IDs ---
const ACCESS_LOG_SS_ID = 'YOUR_HIPAA_ACCESS_LOG_SPREADSHEET_ID'; // Your HIPAA Security Access Log Spreadsheet
const PERSONNEL_SHEET_NAME = 'Personnel'; // This is the name of the sheet (located at the bottom) that has the main personnel info in the google spreadsheet
const ACCESS_LOG_SHEET_NAME = 'Employees'; // This is the name of the sheet (located at the bottom) that has employee names in your Access Log spreadsheet above
const PROVIDER_TRAINING_SHEET_NAME = "Provider Training & Credentialing";
const MEDREC_FOLDER_ID = 'YOUR_MEDREC_FOLDER_ID'; // Main MedRec folder ID (job-classification independent)
const FORMER_EMPLOYEES_FOLDER_ID = 'YOUR_FORMER_EMPLOYEES_FOLDER_ID'; // Former Employees folder ID

// --- Google Doc Template IDs ---
const MED_REC_TEMPLATE_ID = 'YOUR_MED_REC_TEMPLATE_ID'; // Deprecated - See OnboardingHelpers.gs
const HEPB_VAX_FORM_ID = 'YOUR_HEPB_VAX_FORM_ID';    // Hep B Vaccination Form
const ONBOARDING_TEMPLATE_ID = 'YOUR_ONBOARDING_TEMPLATE_ID'; // Onboarding Checklist
const OFFBOARDING_TEMPLATE_ID = 'YOUR_OFFBOARDING_TEMPLATE_ID'; // Offboarding Checklist
const HIPAA_AGREEMNT_ID = 'YOUR_HIPAA_AGREEMENT_ID'; // Deprecated - Handled by form
const OSHA_NEW_TRNG_ID = 'YOUR_OSHA_TRAINING_ID';  // Deprecated - Handled by form
const HANDBOOK_AGRMNT_ID = 'YOUR_HANDBOOK_AGREEMENT_ID'; // Handbook Agreement

// --- Email Template IDs ---
const WELCOME_EMAIL_ID = 'YOUR_WELCOME_EMAIL_TEMPLATE_ID';
const NAMETAG_EMAIL_ID = 'YOUR_NAMETAG_EMAIL_TEMPLATE_ID';
const OFFBOARDING_EMAIL_ID = 'YOUR_OFFBOARDING_EMAIL_TEMPLATE_ID'; // Offboarding Email to Employee
const BEFORE_FIRST_DAY_EMAIL_ID = 'YOUR_BEFORE_FIRST_DAY_EMAIL_TEMPLATE_ID';
const OSHA_TRAINING_ATTESTATION_EMAIL_TEMPLATE_ID = 'YOUR_OSHA_TRAINING_EMAIL_TEMPLATE_ID';
const HIPAA_TRAINING_ATTESTATION_EMAIL_TEMPLATE_ID = 'YOUR_HIPAA_TRAINING_EMAIL_TEMPLATE_ID';

// --- Email Addresses / Groups ---
const PRACTICE_GROUP_EMAIL = 'your-practice-group@yourcompany.com'; // Main practice group for all employees
const NAMETAG_VENDOR_EMAIL = 'vendor@example.com';     // Email for name tag vendor
const RENEWAL_NOTIFICATION_EMAIL = "management@yourcompany.com"; // Recipient for renewal notifications

// --- Spreadsheet Column Headers (Case-sensitive) ---
const COL_FIRST_NAME = "First Name";
const COL_LAST_NAME = "Last Name";
const COL_WORK_EMAIL = "Work Email";
const COL_PERSONAL_EMAIL = "Personal Email";
const COL_START_DATE = "Start Date";
const COL_END_DATE = "End Date"; // Assumed needed for offboarding
const COL_JOB_FOLDER_ID = "Job Folder ID";
// const COL_MED_REC_FOLDER_ID = "Med Rec Folder ID"; // No longer needed - using MEDREC_FOLDER_ID constant instead
const COL_JOB_CLASSIFICATION = "Primary Classification";
const COL_GROUP_EMAIL = "Group Email";
const COL_ACTIVE = "Active";
const COL_EMPLOYEE_DRIVE_FOLDER_ID = "Employee Folder ID"; // Recommend adding this column
const COL_EMPLOYEE_MEDREC_FOLDER_ID = "Employee MedRec Folder ID"; // Recommend adding this column

// --- Other ---
const EMPLOYEE_ACCESS_FOLDER_SUFFIX = 'Employee Access';