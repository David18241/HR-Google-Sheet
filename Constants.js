function myFunction() {
  
}
/**
 * @fileoverview Stores all global constants for the HR Management Script.
 */

// --- Spreadsheet / Folder IDs ---
const ACCESS_LOG_SS_ID = '1l_klIn-rgBd2D26N4JPI7GZXdaDJinm0c127xq7wfzk'; // You HIPAA Security Access Log Spreadsheet
const PERSONNEL_SHEET_NAME = 'Personnel'; // This is the name of the sheet (located at the bottom) that has the main personnel info in the google spreadsheet
const ACCESS_LOG_SHEET_NAME = 'Employees'; // This is the name of the sheet (located at the bottom) that has employee names in your Access Log spreadsheet above
const PROVIDER_TRAINING_SHEET_NAME = "Provider Training & Credentialing";

// --- Google Doc Template IDs ---
const MED_REC_TEMPLATE_ID = '1njsWdsnZ1v-29Do7V1DI3gteok_AMvA_mLoehQXwtuY'; // Deprecated - See OnboardingHelpers.gs
const HEPB_VAX_FORM_ID = '1zz09wN-IZKrG_MwJaLgHvCafdKNYKXG2zpZxJJGKILs';    // Hep B Vaccination Form
const ONBOARDING_TEMPLATE_ID = '1jEdpexG0BqVZ6_4Dq81TRIgr7adfYuWRclbs4q0ZYwA'; // Onboarding Checklist
const OFFBOARDING_TEMPLATE_ID = '1dH_Q-bYzHKvwEgcm3BH6Kzg3cG79bLDS3kE9pF-_Ruo'; // Offboarding Checklist (Assumed - ID was not used in original offboard func)
const HIPAA_AGREEMNT_ID = '1VmTTCN2CM8ep5NlQJvgRYefjz4b-_A9qLY2itpuIeCM'; // Deprecated - Handled by form
const OSHA_NEW_TRNG_ID = '1LQNKkYvj4MNZgYCR8h3kIMH2WV0ugVEPU4Wh2STVj_M';  // Deprecated - Handled by form
const HANDBOOK_AGRMNT_ID = '16wu57hpKPV-bV4Wxam3QNbh2J9kbYWrLjeqTuswBh9Y'; // Handbook Agreement

// --- Email Template IDs ---
const WELCOME_EMAIL_ID = '1CDNt92YZjgmATXQRMCo8LQzC6J71j8EcTNV1rZOrnuk';
const NAMETAG_EMAIL_ID = '1JhfM8j0GigPs3DF3e610qFdKGZGbBa0TZy2dzhfugNg';
const OFFBOARDING_EMAIL_ID = '1Fm9TrN3cDsFJ7m3i1Zl-es7hmBSZ3XJydvfcnUNFwQM'; // Offboarding Email to Employee (Assumed - ID was not used)
const BEFORE_FIRST_DAY_EMAIL_ID = '1FnfbrstcvsmlmyB-g7QzCzj8-iQcw1XTiCPp1oEzLYs';
const OSHA_TRAINING_ATTESTATION_EMAIL_TEMPLATE_ID = '1Bq8rFybGyV3rax4jn-uk1cmNmkTqiKyRc2ZZwJ2KaH0';
const HIPAA_TRAINING_ATTESTATION_EMAIL_TEMPLATE_ID = '1Br3EZJwfJFIdVyUB0RqY7V5iGZm-OAEsiTJHub1m3UY';

// --- Email Addresses / Groups ---
const PRACTICE_GROUP_EMAIL = 'ADCompany@hcoll.org'; // Main practice group
const NAMETAG_VENDOR_EMAIL = 'ma@mickeys.net';     // Name tag vendor
const RENEWAL_NOTIFICATION_EMAIL = "mgmt@hcoll.org"; // Recipient for renewal notifications

// --- Spreadsheet Column Headers (Case-sensitive) ---
const COL_FIRST_NAME = "First Name";
const COL_LAST_NAME = "Last Name";
const COL_WORK_EMAIL = "Work Email";
const COL_PERSONAL_EMAIL = "Personal Email";
const COL_START_DATE = "Start Date";
const COL_END_DATE = "End Date"; // Assumed needed for offboarding
const COL_JOB_FOLDER_ID = "Job Folder ID";
const COL_MED_REC_FOLDER_ID = "Med Rec Folder ID";
const COL_JOB_CLASSIFICATION = "Primary Classification";
const COL_GROUP_EMAIL = "Group Email";
const COL_ACTIVE = "Active";
const COL_EMPLOYEE_DRIVE_FOLDER_ID = "Employee Drive Folder ID"; // Recommend adding this column
const COL_EMPLOYEE_MEDREC_FOLDER_ID = "Employee MedRec Folder ID"; // Recommend adding this column

// --- Other ---
const EMPLOYEE_ACCESS_FOLDER_SUFFIX = 'Employee Access';