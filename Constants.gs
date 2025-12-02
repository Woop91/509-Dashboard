/**
 * ============================================================================
 * CONSTANTS - Centralized Configuration
 * ============================================================================
 *
 * Single source of truth for all constants used throughout the 509 Dashboard.
 * This prevents duplication and makes configuration changes easier.
 *
 * @module Constants
 * @version 2.0.0
 * @author SEIU Local 509 Tech Team
 * ============================================================================
 */

/* ===================== SHEET NAMES ===================== */

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

/* ===================== COLOR SCHEME ===================== */

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

/* ===================== MEMBER DIRECTORY COLUMNS ===================== */

/**
 * Column positions for Member Directory (1-indexed)
 * Use getColumnLetter() to convert to letter notation
 * @const {Object}
 */
const MEMBER_COLS = {
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
  LAST_SURVEY: 16,                 // P
  LAST_EMAIL_OPEN: 17,             // Q
  OPEN_RATE: 18,                   // R
  VOLUNTEER_HOURS: 19,             // S
  INTEREST_LOCAL: 20,              // T
  INTEREST_CHAPTER: 21,            // U
  INTEREST_ALLIED: 22,             // V
  TIMESTAMP: 23,                   // W
  PREFERRED_COMM: 24,              // X
  BEST_TIME: 25,                   // Y
  HAS_OPEN_GRIEVANCE: 26,          // Z
  GRIEVANCE_STATUS: 27,            // AA
  NEXT_DEADLINE: 28,               // AB
  RECENT_CONTACT_DATE: 29,         // AC
  CONTACT_STEWARD: 30,             // AD
  CONTACT_NOTES: 31                // AE
};

/* ===================== GRIEVANCE LOG COLUMNS ===================== */

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

/* ===================== GRIEVANCE TIMELINE CONSTANTS ===================== */

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

/* ===================== GRIEVANCE STATUS VALUES ===================== */

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

/* ===================== CONFIG COLUMN MAPPINGS ===================== */

/**
 * Column mappings for Config sheet
 * @const {Object}
 */
const CONFIG_COLS = {
  JOB_TITLES: 1,           // A
  OFFICE_LOCATIONS: 2,     // B
  UNITS: 3,                // C
  OFFICE_DAYS: 4,          // D
  YES_NO: 5,               // E
  SUPERVISORS: 6,          // F
  MANAGERS: 7,             // G
  STEWARDS: 8,             // H
  GRIEVANCE_STATUS: 9,     // I
  GRIEVANCE_STEP: 10,      // J
  ISSUE_CATEGORY: 11,      // K
  ARTICLES_VIOLATED: 12,   // L
  COMM_METHODS: 13,        // M
  COORDINATOR_1_NAME: 16,  // P
  COORDINATOR_2_NAME: 17,  // Q
  COORDINATOR_3_NAME: 18,  // R
  COORDINATOR_1_EMAIL: 19, // S
  COORDINATOR_2_EMAIL: 20, // T
  COORDINATOR_3_EMAIL: 21  // U
};

/* ===================== CACHE CONFIGURATION ===================== */

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
  ANALYTICS_DATA: 'analytics_data'
};

/* ===================== ERROR CONFIGURATION ===================== */

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

/* ===================== UI CONFIGURATION ===================== */

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

/* ===================== EMAIL CONFIGURATION ===================== */

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

/* ===================== PERFORMANCE CONFIGURATION ===================== */

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

/* ===================== FEATURE FLAGS ===================== */

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

/* ===================== ERROR MESSAGES ===================== */

/**
 * Standardized error and notification messages
 * Ensures consistent UX across the application
 * @const {Object}
 */
const ERROR_MESSAGES = {
  // Access and permissions
  ACCESS_DENIED: (operation) => `⛔ Access Denied: ${operation}`,
  INSUFFICIENT_ROLE: (required, current) => `⛔ Insufficient Permissions: ${required} role required (you have ${current})`,

  // Not found errors
  NOT_FOUND: (item) => `❌ Not Found: ${item}`,
  SHEET_NOT_FOUND: (sheetName) => `❌ Sheet Not Found: ${sheetName}`,
  MEMBER_NOT_FOUND: (memberId) => `❌ Member Not Found: ${memberId}`,
  GRIEVANCE_NOT_FOUND: (grievanceId) => `❌ Grievance Not Found: ${grievanceId}`,

  // Validation errors
  INVALID_INPUT: (field) => `⚠️ Invalid Input: ${field}`,
  INVALID_EMAIL: 'Invalid email address format',
  INVALID_PHONE: 'Invalid phone number format',
  INVALID_DATE: 'Invalid date format',
  REQUIRED_FIELD: (field) => `⚠️ Required Field: ${field} cannot be empty`,

  // Rate limiting
  RATE_LIMIT: (seconds) => `⏱️ Rate Limit Exceeded: Please wait ${seconds} seconds before trying again`,

  // Success messages
  SUCCESS: (action) => `✅ Success: ${action}`,
  CREATED: (item) => `✅ Created: ${item}`,
  UPDATED: (item) => `✅ Updated: ${item}`,
  DELETED: (item) => `✅ Deleted: ${item}`,

  // Progress messages
  IN_PROGRESS: (action) => `⏳ ${action}...`,
  LOADING: (item) => `⏳ Loading ${item}...`,
  PROCESSING: (item) => `⏳ Processing ${item}...`,

  // Warnings
  WARNING: (message) => `⚠️ Warning: ${message}`,
  DATA_LOSS_WARNING: 'This action cannot be undone. Are you sure?',

  // System errors
  SYSTEM_ERROR: 'An unexpected error occurred. Please try again.',
  NETWORK_ERROR: 'Network error. Please check your connection.',
  TIMEOUT_ERROR: 'Operation timed out. Please try again.',

  // Data integrity
  DATA_INTEGRITY_ERROR: 'Data integrity check failed',
  DUPLICATE_ENTRY: (field) => `⚠️ Duplicate Entry: ${field} already exists`
};

/* ===================== ADMIN CONFIGURATION ===================== */

/**
 * Admin email configuration
 * Note: For security, admin emails should also be configured in the Config sheet.
 * These are fallback defaults if Config sheet is not set up yet.
 * @const {Object}
 */
const ADMIN_CONFIG = {
  // Column position in Config sheet for admin emails
  CONFIG_COLUMN: 19, // Column S - Admin Emails

  // Fallback admin emails (used only if Config sheet not available)
  FALLBACK_ADMINS: [
    'admin@seiu509.org',
    'president@seiu509.org',
    'techsupport@seiu509.org'
  ],

  // Admin permissions
  CAN_MODIFY_CONFIG: true,
  CAN_VIEW_AUDIT_LOG: true,
  CAN_RUN_SECURITY_AUDIT: true,
  CAN_MANAGE_USERS: true,
  CAN_DELETE_DATA: true,
  CAN_EXPORT_DATA: true
};

/* ===================== AUDIT LOG CONFIGURATION ===================== */

/**
 * Audit log settings and constants
 * @const {Object}
 */
const AUDIT_LOG_CONFIG = {
  MAX_ENTRIES: 10000,          // Maximum number of audit log entries to keep
  HEADER_ROWS: 1,              // Number of header rows in audit log
  AUTO_TRIM_ENABLED: true,     // Automatically trim old entries when limit reached
  RETENTION_DAYS: 365,         // Days to retain audit entries
  LOG_SHEET_NAME: '🔒 Audit Log'
};

/* ===================== MAGIC NUMBER REPLACEMENTS ===================== */

/**
 * Named constants for commonly used numeric values
 * Replaces "magic numbers" throughout the codebase for clarity
 * @const {Object}
 */
const NUMERIC_CONSTANTS = {
  // Cache TTL (Time To Live) in seconds
  CACHE_TTL_SHORT: 300,        // 5 minutes
  CACHE_TTL_MEDIUM: 3600,      // 1 hour
  CACHE_TTL_LONG: 21600,       // 6 hours
  CACHE_TTL_VERY_LONG: 86400,  // 24 hours

  // Performance thresholds in milliseconds
  SLOW_OPERATION_MS: 60000,    // 1 minute
  VERY_SLOW_MS: 120000,        // 2 minutes
  MAX_EXECUTION_TIME_MS: 300000, // 5 minutes (Apps Script limit: 6 minutes)

  // Data limits
  MAX_BATCH_SIZE: 1000,        // Maximum rows per batch operation
  MAX_SEARCH_RESULTS: 100,     // Maximum search results to display at once
  LAZY_LOAD_THRESHOLD: 5000,   // Start lazy loading after this many rows

  // Time intervals in milliseconds
  DEBOUNCE_DELAY_MS: 300,      // Debounce delay for search/input
  TOAST_DURATION_SHORT: 2,     // Short toast notification (seconds)
  TOAST_DURATION_MEDIUM: 5,    // Medium toast notification (seconds)
  TOAST_DURATION_LONG: 10,     // Long toast notification (seconds)

  // Archive settings
  ARCHIVE_AFTER_YEARS: 2,      // Archive grievances closed for 2+ years
  ARCHIVE_BATCH_SIZE: 500,     // Archive in batches of 500

  // Pagination
  ITEMS_PER_PAGE: 100,         // Items to display per page
  MAX_PAGES_TO_SHOW: 10        // Maximum page numbers to show in pagination
};

/* ===================== VERSION INFORMATION ===================== */

/**
 * Version information
 * @const {Object}
 */
const VERSION_INFO = {
  MAJOR: 2,
  MINOR: 1,
  PATCH: 0,
  BUILD: '20251202-improvements',
  RELEASE_DATE: '2025-12-02',
  CODENAME: 'Security Enhanced + Code Review Improvements'
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

/* ===================== UTILITY FUNCTIONS ===================== */

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
  let letter = '';
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
  let number = 0;
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
  requiredSheets.forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      missing.push(sheetName);
    }
  });

  return {
    valid: missing.length === 0,
    missing: missing
  };
}
