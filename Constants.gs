/**
 * ------------------------------------------------------------------------====
 * CONSTANTS - Centralized Configuration
 * ------------------------------------------------------------------------====
 *
 * Single source of truth for all constants used throughout the 509 Dashboard.
 * This prevents duplication and makes configuration changes easier.
 *
 * @module Constants
 * @version 2.0.0
 * @author SEIU Local 509 Tech Team
 * ------------------------------------------------------------------------====
 */

/* --------------------= SHEET NAMES --------------------= */

/**
 * Sheet names used throughout the application
 * @const {Object}
 */
const SHEETS = {
  CONFIG: "Config",
  MEMBER_DIR: "Member Directory",
  GRIEVANCE_LOG: "Grievance Log",
  DASHBOARD: "Dashboard",
  ANALYTICS: "Analytics Data",
  FEEDBACK: "Feedback & Development",
  MEMBER_SATISFACTION: "Member Satisfaction",
  INTERACTIVE_DASHBOARD: "🎯 Interactive (Your Custom View)",
  STEWARD_WORKLOAD: "👨‍⚖️ Steward Workload",
  TRENDS: "📈 Trends & Timeline",
  PERFORMANCE: "🎯 Test 2: Performance",
  LOCATION: "🗺️ Location Analytics",
  TYPE_ANALYSIS: "📊 Type Analysis",
  EXECUTIVE_DASHBOARD: "💼 Executive Dashboard",
  EXECUTIVE: "💼 Executive Dashboard",  // Alias for backward compatibility
  KPI_PERFORMANCE: "📊 KPI Performance Dashboard",
  KPI_BOARD: "📊 KPI Performance Dashboard",  // Alias for backward compatibility
  MEMBER_ENGAGEMENT: "👥 Member Engagement",
  COST_IMPACT: "💰 Cost Impact",
  QUICK_STATS: "⚡ Quick Stats",
  ARCHIVE: "📦 Archive",
  DIAGNOSTICS: "🔧 Diagnostics",
  GETTING_STARTED: "🚀 Getting Started",
  FAQ: "❓ FAQ & Help",
  USER_SETTINGS: "⚙️ User Settings"
};

/* --------------------= COLOR SCHEME --------------------= */

/**
 * Color palette for consistent theming
 * @const {Object}
 */
const COLORS = {
  // Primary brand colors
  PRIMARY_BLUE: "#7EC8E3",
  PRIMARY_PURPLE: "#7C3AED",
  UNION_GREEN: "#059669",
  SOLIDARITY_RED: "#DC2626",

  // Accent colors
  ACCENT_TEAL: "#14B8A6",
  ACCENT_PURPLE: "#7C3AED",
  ACCENT_ORANGE: "#F97316",
  ACCENT_YELLOW: "#FCD34D",

  // Neutral colors
  WHITE: "#FFFFFF",
  LIGHT_GRAY: "#F3F4F6",
  BORDER_GRAY: "#D1D5DB",
  TEXT_GRAY: "#6B7280",
  TEXT_DARK: "#1F2937",
  BLACK: "#000000",

  // Specialty colors
  CARD_BG: "#FAFAFA",
  INFO_LIGHT: "#E0E7FF",
  SUCCESS_LIGHT: "#D1FAE5",
  WARNING_LIGHT: "#FEF3C7",
  ERROR_LIGHT: "#FEE2E2",
  HEADER_BLUE: "#3B82F6",
  HEADER_GREEN: "#10B981",

  // Status colors
  STATUS_OPEN: "#DC2626",        // Red
  STATUS_PENDING: "#F97316",     // Orange
  STATUS_SETTLED: "#059669",     // Green
  STATUS_CLOSED: "#6B7280",      // Gray
  STATUS_APPEALED: "#7C3AED"     // Purple
};

/* --------------------= MEMBER DIRECTORY COLUMNS --------------------= */

/**
 * Column positions for Member Directory (1-indexed)
 * Use getColumnLetter() to convert to letter notation
 * @const {Object}
 */
const MEMBER_COLS = {
  // 27 columns total (removed 5 unused: LAST_SURVEY, LAST_EMAIL_OPEN, TIMESTAMP, PREFERRED_COMM, BEST_TIME)
  MEMBER_ID: 1,                    // A
  FIRST_NAME: 2,                   // B
  LAST_NAME: 3,                    // C
  JOB_TITLE: 4,                    // D
  WORK_LOCATION: 5,                // E
  UNIT: 6,                         // F
  OFFICE_DAYS: 7,                  // G
  EMAIL: 8,                        // H
  PHONE: 9,                        // I
  IS_STEWARD: 10,                  // J
  SUPERVISOR: 11,                  // K
  MANAGER: 12,                     // L
  ASSIGNED_STEWARD: 13,            // M
  LAST_VIRTUAL_MTG: 14,            // N
  LAST_INPERSON_MTG: 15,           // O
  OPEN_RATE: 16,                   // P
  VOLUNTEER_HOURS: 17,             // Q
  INTEREST_LOCAL: 18,              // R
  INTEREST_CHAPTER: 19,            // S
  INTEREST_ALLIED: 20,             // T
  HAS_OPEN_GRIEVANCE: 21,          // U
  GRIEVANCE_STATUS: 22,            // V
  NEXT_DEADLINE: 23,               // W
  RECENT_CONTACT_DATE: 24,         // X
  CONTACT_STEWARD: 25,             // Y
  CONTACT_NOTES: 26,               // Z
  START_GRIEVANCE: 27              // AA - Checkbox to start grievance with prepopulated member info
};

/* --------------------= GRIEVANCE LOG COLUMNS --------------------= */

/**
 * Column positions for Grievance Log (1-indexed)
 * Use getColumnLetter() to convert to letter notation
 * @const {Object}
 */
const GRIEVANCE_COLS = {
  GRIEVANCE_ID: 1,      // A
  MEMBER_ID: 2,         // B
  FIRST_NAME: 3,        // C
  LAST_NAME: 4,         // D
  STATUS: 5,            // E
  CURRENT_STEP: 6,      // F
  INCIDENT_DATE: 7,     // G
  FILING_DEADLINE: 8,   // H
  DATE_FILED: 9,        // I
  STEP1_DUE: 10,        // J
  STEP1_RCVD: 11,       // K
  STEP2_APPEAL_DUE: 12, // L
  STEP2_APPEAL_FILED: 13, // M
  STEP2_DUE: 14,        // N
  STEP2_RCVD: 15,       // O
  STEP3_APPEAL_DUE: 16, // P
  STEP3_APPEAL_FILED: 17, // Q
  DATE_CLOSED: 18,      // R
  DAYS_OPEN: 19,        // S
  NEXT_ACTION_DUE: 20,  // T
  DAYS_TO_DEADLINE: 21, // U
  ARTICLES: 22,         // V
  ISSUE_CATEGORY: 23,   // W
  MEMBER_EMAIL: 24,     // X
  UNIT: 25,             // Y
  LOCATION: 26,         // Z
  STEWARD: 27,          // AA
  RESOLUTION: 28        // AB
};

/* --------------------= GRIEVANCE TIMELINE CONSTANTS --------------------= */

/**
 * Timeline rules for grievance deadlines (in days)
 * Based on standard union contract
 * @const {Object}
 */
const GRIEVANCE_TIMELINES = {
  FILING_DEADLINE_DAYS: 21,      // Days after incident to file
  STEP1_DECISION_DAYS: 30,       // Days for Step I decision
  STEP2_APPEAL_DAYS: 10,         // Days to appeal to Step II
  STEP2_DECISION_DAYS: 30,       // Days for Step II decision
  STEP3_APPEAL_DAYS: 30,         // Days to appeal to Step III
  STEP3_DECISION_DAYS: 45,       // Days for Step III decision
  ARBITRATION_FILING_DAYS: 30,   // Days to file for arbitration
  WARNING_THRESHOLD_DAYS: 7      // Show warning when deadline within X days
};

/* --------------------= GRIEVANCE STATUS VALUES --------------------= */

/**
 * Valid grievance status values
 * @const {Array<string>}
 */
const GRIEVANCE_STATUSES = [
  'Open',
  'Pending Info',
  'Settled',
  'Withdrawn',
  'Closed',
  'Appealed',
  'In Arbitration',
  'Denied'
];

/**
 * Valid grievance step values
 * @const {Array<string>}
 */
const GRIEVANCE_STEPS = [
  'Informal',
  'Step I',
  'Step II',
  'Step III',
  'Mediation',
  'Arbitration'
];

/**
 * Valid issue categories
 * @const {Array<string>}
 */
const ISSUE_CATEGORIES = [
  'Discipline',
  'Workload',
  'Scheduling',
  'Pay/Compensation',
  'Discrimination',
  'Safety',
  'Benefits',
  'Harassment',
  'Performance Evaluation',
  'Job Classification',
  'Layoff/Recall',
  'Leave',
  'Other'
];

/* --------------------= CONFIG COLUMN MAPPINGS --------------------= */

/**
 * Column mappings for Config sheet
 * @const {Object}
 */
const CONFIG_COLS = {
  // Employment Info (cols 1-5)
  JOB_TITLES: 1,              // A
  OFFICE_LOCATIONS: 2,        // B
  UNITS: 3,                   // C
  OFFICE_DAYS: 4,             // D
  YES_NO: 5,                  // E
  // Supervision (cols 6-8) - combined first/last names
  SUPERVISORS: 6,             // F - full name (First Last)
  MANAGERS: 7,                // G - full name (First Last)
  STEWARDS: 8,                // H
  // Grievance Settings (cols 9-13)
  GRIEVANCE_STATUS: 9,        // I
  GRIEVANCE_STEP: 10,         // J
  ISSUE_CATEGORY: 11,         // K
  ARTICLES_VIOLATED: 12,      // L
  COMM_METHODS: 13,           // M
  // Links & Coordinators (cols 14-16)
  GRIEVANCE_COORDINATORS: 14, // N - comma-separated list
  GRIEVANCE_FORM_URL: 15,     // O - URL to grievance intake form
  CONTACT_FORM_URL: 16,       // P - URL to contact form
  // Notifications (cols 17-19)
  ADMIN_EMAILS: 17,           // Q - admin email addresses
  ALERT_DAYS: 18,             // R - days before deadline to alert (e.g., "3, 7, 14")
  NOTIFICATION_RECIPIENTS: 19,// S - default CC for notifications
  // Organization (cols 20-23)
  ORG_NAME: 20,               // T - organization name
  LOCAL_NUMBER: 21,           // U - union local number
  MAIN_ADDRESS: 22,           // V - main office address
  MAIN_PHONE: 23,             // W - main phone number
  // Integration (cols 24-25)
  DRIVE_FOLDER_ID: 24,        // X - Google Drive root folder ID
  CALENDAR_ID: 25,            // Y - Google Calendar ID
  // Deadlines (cols 26-29)
  FILING_DEADLINE_DAYS: 26,   // Z - days to file grievance (default: 21)
  STEP1_RESPONSE_DAYS: 27,    // AA - days for Step I response (default: 30)
  STEP2_APPEAL_DAYS: 28,      // AB - days to appeal to Step II (default: 10)
  STEP2_RESPONSE_DAYS: 29     // AC - days for Step II response (default: 30)
};

/* --------------------= CACHE CONFIGURATION --------------------= */

/**
 * Cache configuration for performance optimization
 * @const {Object}
 */
const CACHE_CONFIG = {
  MEMORY_TTL: 300,              // 5 minutes (in seconds)
  PROPERTIES_TTL: 3600,         // 1 hour (in seconds)
  DOCUMENT_TTL: 21600,          // 6 hours (in seconds)
  MAX_CACHE_SIZE_BYTES: 100000, // 100KB max per cache entry
  ENABLED: true                 // Global cache enable/disable
};

/**
 * Cache keys used throughout the application
 * @const {Object}
 */
const CACHE_KEYS = {
  MEMBER_COUNT: 'member_count',
  GRIEVANCE_COUNT: 'grievance_count',
  STEWARD_WORKLOAD: 'steward_workload',
  MEMBER_LIST: 'member_list',
  CONFIG_DATA: 'config_data',
  DASHBOARD_METRICS: 'dashboard_metrics',
  ANALYTICS_DATA: 'analytics_data',
  // Additional keys for data caching layer
  ALL_GRIEVANCES: 'all_grievances',
  ALL_MEMBERS: 'all_members',
  ALL_STEWARDS: 'all_stewards'
};

/* --------------------= ERROR CONFIGURATION --------------------= */

/**
 * Error handling and logging configuration
 * @const {Object}
 */
const ERROR_CONFIG = {
  LOG_SHEET_NAME: 'Error_Log',
  MAX_LOG_ENTRIES: 1000,
  ERROR_LEVELS: {
    INFO: 'INFO',
    WARNING: 'WARNING',
    ERROR: 'ERROR',
    CRITICAL: 'CRITICAL'
  },
  NOTIFICATION_THRESHOLD: 'ERROR',
  AUTO_RECOVERY_ENABLED: true,
  SHOW_USER_ERRORS: true,
  LOG_TO_CONSOLE: true
};

/**
 * Error categories for classification
 * @const {Object}
 */
const ERROR_CATEGORIES = {
  VALIDATION: 'VALIDATION',
  PERMISSION: 'PERMISSION',
  NETWORK: 'NETWORK',
  DATA_INTEGRITY: 'DATA_INTEGRITY',
  USER_INPUT: 'USER_INPUT',
  SYSTEM: 'SYSTEM',
  INTEGRATION: 'INTEGRATION',
  SECURITY: 'SECURITY'
};

/* --------------------= UI CONFIGURATION --------------------= */

/**
 * User interface configuration
 * @const {Object}
 */
const UI_CONFIG = {
  TOAST_DURATION: 3,          // Seconds
  DIALOG_WIDTH: 700,          // Pixels
  DIALOG_HEIGHT: 500,         // Pixels
  MAX_DROPDOWN_ITEMS: 500,    // Max items in dropdown before search
  PAGINATION_SIZE: 100,       // Rows per page
  DEFAULT_THEME: 'light'      // light or dark
};

/* --------------------= EMAIL CONFIGURATION --------------------= */

/**
 * Email configuration
 * @const {Object}
 */
const EMAIL_CONFIG = {
  FROM_NAME: 'SEIU Local 509',
  REPLY_TO: 'info@seiu509.org',
  GRIEVANCE_EMAIL: 'grievances@seiu509.org',
  MAX_SUBJECT_LENGTH: 200,
  MAX_BODY_LENGTH: 10000,
  ATTACHMENTS_ENABLED: true,
  MAX_ATTACHMENT_SIZE_MB: 25
};

/* --------------------= PERFORMANCE CONFIGURATION --------------------= */

/**
 * Performance and batch operation settings
 * @const {Object}
 */
const PERFORMANCE_CONFIG = {
  BATCH_SIZE: 1000,           // Rows per batch for bulk operations
  MAX_EXECUTION_TIME: 300,    // Max execution time in seconds (5 min)
  LAZY_LOAD_THRESHOLD: 5000,  // Load data lazily after this many rows
  ENABLE_PROFILING: false,    // Performance profiling
  AUTO_FLUSH_ENABLED: true    // Auto SpreadsheetApp.flush()
};

/* --------------------= FEATURE FLAGS --------------------= */

/**
 * Feature flags for enabling/disabling features
 * @const {Object}
 */
const FEATURE_FLAGS = {
  ENABLE_DARK_MODE: true,
  ENABLE_KEYBOARD_SHORTCUTS: true,
  ENABLE_MOBILE_OPTIMIZATION: true,
  ENABLE_PREDICTIVE_ANALYTICS: true,
  ENABLE_AUTO_ASSIGNMENT: true,
  ENABLE_CALENDAR_INTEGRATION: true,
  ENABLE_GMAIL_INTEGRATION: true,
  ENABLE_DRIVE_INTEGRATION: true,
  ENABLE_ADVANCED_SEARCH: true,
  ENABLE_CUSTOM_REPORTS: true,
  ENABLE_DATA_EXPORT: true,
  ENABLE_UNDO_REDO: true,
  ENABLE_ADHD_FEATURES: true
};

/* --------------------= VERSION INFORMATION --------------------= */

/**
 * Version information
 * @const {Object}
 */
const VERSION_INFO = {
  MAJOR: 2,
  MINOR: 0,
  PATCH: 0,
  BUILD: '20251202',
  RELEASE_DATE: '2025-12-02',
  CODENAME: 'Security Enhanced'
};

/**
 * Gets version string
 * @returns {string} Version string (e.g., "2.0.0")
 */
function getVersionString() {
  return `${VERSION_INFO.MAJOR}.${VERSION_INFO.MINOR}.${VERSION_INFO.PATCH}`;
}

/**
 * Gets full version info string
 * @returns {string} Full version string with build info
 */
function getFullVersionString() {
  return `v${getVersionString()} (${VERSION_INFO.CODENAME}) - Build ${VERSION_INFO.BUILD}`;
}

/* --------------------= UTILITY FUNCTIONS --------------------= */

/**
 * Converts a column number to letter notation (1=A, 27=AA, etc.)
 * @param {number} columnNumber - Column number (1-indexed)
 * @returns {string} Column letter (A, B, AA, etc.)
 *
 * @example
 * getColumnLetter(1)  // Returns "A"
 * getColumnLetter(27) // Returns "AA"
 */
function getColumnLetter(columnNumber) {
  var letter = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}

/**
 * Converts column letter to number (A=1, AA=27, etc.)
 * @param {string} columnLetter - Column letter
 * @returns {number} Column number (1-indexed)
 *
 * @example
 * getColumnNumber('A')  // Returns 1
 * getColumnNumber('AA') // Returns 27
 */
function getColumnNumber(columnLetter) {
  var number = 0;
  for (let i = 0; i < columnLetter.length; i++) {
    number = number * 26 + (columnLetter.charCodeAt(i) - 64);
  }
  return number;
}

/**
 * Gets column letter for a named column in Member Directory
 * @param {string} columnName - Column name from MEMBER_COLS
 * @returns {string} Column letter
 *
 * @example
 * getMemberColumn('EMAIL') // Returns "H"
 */
function getMemberColumn(columnName) {
  if (!MEMBER_COLS[columnName]) {
    throw new Error(`Unknown member column: ${columnName}`);
  }
  return getColumnLetter(MEMBER_COLS[columnName]);
}

/**
 * Gets column letter for a named column in Grievance Log
 * @param {string} columnName - Column name from GRIEVANCE_COLS
 * @returns {string} Column letter
 *
 * @example
 * getGrievanceColumn('STATUS') // Returns "E"
 */
function getGrievanceColumn(columnName) {
  if (!GRIEVANCE_COLS[columnName]) {
    throw new Error(`Unknown grievance column: ${columnName}`);
  }
  return getColumnLetter(GRIEVANCE_COLS[columnName]);
}

/**
 * Validates that all required sheets exist
 * @returns {Object} {valid: boolean, missing: Array<string>}
 */
function validateRequiredSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    SHEETS.CONFIG,
    SHEETS.MEMBER_DIR,
    SHEETS.GRIEVANCE_LOG,
    SHEETS.DASHBOARD
  ];

  const missing = [];
  requiredSheets.forEach(function(sheetName) {
    if (!ss.getSheetByName(sheetName)) {
      missing.push(sheetName);
    }
  });

  return {
    valid: missing.length === 0,
    missing: missing
  };
}

/* --------------------= CONFIGURATION REFERENCE --------------------= */

/**
 * CONFIGURATION REFERENCE
 * =======================
 *
 * This section documents where all configuration is located for easy reference.
 * All config should be defined ONCE and only in its designated module.
 *
 * SHEETS, COLORS, COLUMNS (this file - Constants.gs):
 *   - SHEETS: Sheet name constants
 *   - COLORS: Color palette
 *   - MEMBER_COLS: Member Directory column positions
 *   - GRIEVANCE_COLS: Grievance Log column positions
 *   - CONFIG_COLS: Config sheet column positions
 *   - GRIEVANCE_TIMELINES: Timeline rules for grievances
 *   - GRIEVANCE_STATUSES: Valid status values
 *   - GRIEVANCE_STEPS: Valid step values
 *   - ISSUE_CATEGORIES: Valid issue categories
 *
 * CACHE & PERFORMANCE (this file - Constants.gs):
 *   - CACHE_CONFIG: Cache TTL and size limits
 *   - CACHE_KEYS: Cache key constants
 *   - PERFORMANCE_CONFIG: Batch sizes, execution limits
 *
 * UI & EMAIL (this file - Constants.gs):
 *   - UI_CONFIG: Dialog sizes, pagination, theme defaults
 *   - EMAIL_CONFIG: Email settings and limits
 *
 * ERROR HANDLING (this file - Constants.gs):
 *   - ERROR_CONFIG: Error logging configuration
 *   - ERROR_CATEGORIES: Error classification
 *
 * FEATURE FLAGS (this file - Constants.gs):
 *   - FEATURE_FLAGS: Enable/disable features
 *   - VERSION_INFO: Version numbers and build info
 *
 * SECURITY & ROLES (SecurityUtils.gs):
 *   - SECURITY_ROLES: Simple role enum (ADMIN, STEWARD, MEMBER, GUEST)
 *   - ADMIN_EMAILS: List of admin email addresses
 *   - AUDIT_LOG_SHEET: Audit log sheet name
 *   - RATE_LIMITS: Rate limiting configuration
 *
 * ADVANCED ROLES (SecurityService.gs):
 *   - ROLES: Detailed role definitions with permissions
 *     Includes: ADMIN, STEWARD, COORDINATOR, MEMBER, VIEWER
 *     Each role has: name, permissions[], description
 *
 * HOW TO UPDATE CONFIGURATION:
 *   1. Find the constant in the appropriate module above
 *   2. Edit ONLY in that module (never duplicate)
 *   3. Run: node build.js
 *   4. The build will fail if duplicates are detected
 *
 * HOW TO ADD ADMIN USERS:
 *   1. Edit SecurityUtils.gs
 *   2. Add email to ADMIN_EMAILS array
 *   3. Run: node build.js
 *
 * @see SecurityUtils.gs for role/permission functions
 * @see SecurityService.gs for advanced RBAC
 */
