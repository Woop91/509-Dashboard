# TODO / Feature Roadmap

Future feature ideas and enhancements for the 509 Dashboard.

**Status Summary:** 27 of 48 original features have been implemented. 21 remain pending.

---

## Pending Features (21 items)

### HIGH PRIORITY (8 items)

#### 1. Email Unsubscribe / Opt-Out System
**Complexity:** Medium | **Status:** Planned

Add the ability to mark members as having opted out of email communications:
- Checkbox column in Member Directory
- Light red row highlighting when opted out (`#FFCDD2`)
- Export prefixes email with `(UNSUBSCRIBED)` to prevent accidental sends
- Filter opted-out members from bulk emails

**Files:** Constants.gs, MemberDirectory.gs, ReportGeneration.gs, EmailIntegration.gs

---

#### 2. Extend Auto-Formula Coverage
**Complexity:** Simple | **Status:** Planned

Extend formulas from 100 rows to 1000+ using ARRAYFORMULA.
- Replace individual `setFormula` calls with ARRAYFORMULA implementations

---

#### 3. Optimize Seed Data Performance
**Complexity:** Moderate | **Status:** Planned

Improve batch processing for 20k member seeding (currently 2-3 minutes).
- Use `SpreadsheetApp.flush()` strategically
- Add progress toasts every 500 rows

---

#### 4. Create Quick Actions Menu
**Complexity:** Moderate | **Status:** Planned

Add right-click context menu for common actions:
- Start Grievance, View History, Email
- Implement onSelectionChange trigger with context-aware options

---

#### 5. Implement Change Tracking
**Complexity:** Moderate | **Status:** Planned

Track all data modifications with audit trail:
- Timestamp, user, field, old/new values
- Create onEdit trigger to log changes to Change Log sheet

---

#### 6. Trend Analysis & Forecasting
**Complexity:** Moderate | **Status:** Planned

Identify patterns over time:
- Month-over-month comparisons
- Seasonal patterns
- Volume forecasting and early warning system

---

#### 7. PII Protection
**Complexity:** Complex | **Status:** Planned

Protect sensitive member data:
- Encrypt sensitive fields
- Mask data in exports
- GDPR/CCPA compliance

---

#### 8. Interactive Tutorial
**Complexity:** Moderate | **Status:** Planned

In-app guided tour for first-time users:
- Walkthrough with highlighted features
- Interactive steps

---

### MEDIUM PRIORITY (3 items)

#### 9. Optimize QUERY Formulas
**Complexity:** Simple | **Status:** Planned

Consolidate multiple QUERY calls to reduce calculation time.
- Use virtual tables and GROUP BY instead of multiple COUNTIF calls

---

#### 10. Slack/Teams Integration
**Complexity:** Moderate | **Status:** Planned

Real-time notifications with bot commands:
- Post updates to channels
- Notify on deadlines
- Allow status updates from chat

---

#### 11. Context-Sensitive Help
**Complexity:** Simple | **Status:** Planned

Help button on each sheet:
- ? icon explaining purpose
- List common tasks

---

### LOW PRIORITY (10 items)

#### 12. Template System
**Complexity:** Moderate | **Status:** Planned

Pre-built grievance templates for common issue types with auto-fill.

---

#### 13. Benchmark Comparisons
**Complexity:** Moderate | **Status:** Planned

Compare performance against industry standards with percentile rankings.

---

#### 14. Real-Time Dashboard Updates
**Complexity:** Moderate | **Status:** Planned

Live data refresh with auto-update every 5 minutes via triggers.

---

#### 15. API Layer
**Complexity:** Complex | **Status:** Planned

RESTful API for external access:
- Google Apps Script Web App
- GET/POST/PUT endpoints
- API key authentication

---

#### 16. Zapier/Make.com Integration
**Complexity:** Moderate | **Status:** Planned

Connect to 1000+ apps via webhooks.

---

#### 17. Progressive Web App (PWA)
**Complexity:** Complex | **Status:** Planned

Install as app on mobile devices for native experience.

---

#### 18. Offline Mode
**Complexity:** Very Complex | **Status:** Planned

Work without internet:
- Local storage cache
- Sync when online
- Conflict resolution

---

#### 19. SMS Notifications
**Complexity:** Moderate | **Status:** Planned

Text alerts for critical items via Twilio integration.

---

#### 20. Session Management
**Complexity:** Moderate | **Status:** Planned

Track active users:
- Show current viewers
- Lock rows being edited
- Session timeout

---

#### 21. Release Notes
**Complexity:** Simple | **Status:** Planned

Track version changes:
- CHANGELOG sheet
- Auto-notify on updates
- Highlight new features

---

## Implemented Features (27 items)

The following features have been completed:

| Category | Feature | Implementation |
|----------|---------|----------------|
| **Performance** | Data Pagination | DataPagination.gs |
| | Caching Layer | DataCachingLayer.gs |
| **UX** | Member Search | MemberSearch.gs |
| | Keyboard Shortcuts | KeyboardShortcuts.gs |
| | Undo/Redo | UndoRedoSystem.gs |
| | ADHD Features | EnhancedADHDFeatures.gs |
| | Mobile Views | MobileOptimization.gs |
| | Dark Mode | DarkModeThemes.gs |
| **Data Integrity** | Email/Phone Validation | DataIntegrityEnhancements.gs |
| | Duplicate Detection | DataIntegrityEnhancements.gs |
| | Data Quality Dashboard | DataIntegrityEnhancements.gs |
| | Referential Integrity | DataIntegrityEnhancements.gs |
| **Automation** | Deadline Notifications | AutomatedNotifications.gs |
| | Batch Operations | BatchOperations.gs |
| | Automated Reports | AutomatedReports.gs |
| | Workflow State Machine | WorkflowStateMachine.gs |
| | Smart Auto-Assignment | SmartAutoAssignment.gs |
| **Analytics** | Predictive Analytics | PredictiveAnalytics.gs |
| | Custom Report Builder | CustomReportBuilder.gs |
| | Root Cause Analysis | RootCauseAnalysis.gs |
| **Integration** | Google Calendar | CalendarIntegration.gs |
| | Gmail | GmailIntegration.gs |
| | Google Drive | GoogleDriveIntegration.gs |
| **Security** | RBAC | SecurityService.gs |
| | Audit Logging | AuditLoggingRBAC.gs |
| | Data Backup | DataBackupRecovery.gs |
| **Docs** | FAQ Database | FAQKnowledgeBase.gs |

---

## Notes

- Feature requests can be added to this file
- Prioritize based on user impact and implementation complexity
- Update status as work progresses: Planned -> In Progress -> Completed
- Video Tutorials (High priority) requires external recording - not a code task
