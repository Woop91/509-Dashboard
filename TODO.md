# TODO / Feature Roadmap

Future feature ideas and enhancements for the 509 Dashboard.

---

## Pending Features

### 1. Email Unsubscribe / Opt-Out System

**Priority:** Medium
**Complexity:** Medium
**Status:** Planned

#### Description
Add the ability to mark members as having opted out of email communications, with visual indicators and export safeguards.

#### Requirements

1. **New Column in Member Directory**
   - Add "Email Opt-Out" checkbox column
   - Position: After Email column or in a logical grouping
   - Data validation: Checkbox (TRUE/FALSE)

2. **Visual Indicator - Row Highlighting**
   - When "Email Opt-Out" is checked (TRUE), the entire row background turns light red
   - Suggested color: `#FFCDD2` (Material Design Red 100) or `#FFEBEE` (Red 50)
   - Implementation: Use conditional formatting or onEdit trigger

3. **Export Safeguard**
   - When exporting member data (CSV, reports, etc.), modify opted-out email addresses
   - Prefix email with `(UNSUBSCRIBED)` - e.g., `(UNSUBSCRIBED)john.doe@email.com`
   - This ensures that if someone ignores the visual warning and attempts to use the email, the send will fail

4. **Email Composition Integration**
   - Filter out opted-out members from bulk email recipient lists
   - Show warning if manually attempting to email an opted-out member
   - Add visual indicator in email compose dialogs

#### Implementation Notes

**Column Addition (Constants.gs):**
```javascript
// Add to MEMBER_COLS
EMAIL_OPT_OUT: XX,  // New column number
```

**Conditional Formatting (MemberDirectory.gs or Setup):**
```javascript
// Apply light red background when Email Opt-Out is TRUE
const rule = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('=$XX2=TRUE')  // XX = opt-out column
  .setBackground('#FFCDD2')
  .setRanges([memberRange])
  .build();
```

**Export Modification (ReportGeneration.gs):**
```javascript
// In export functions, check opt-out status
if (row[MEMBER_COLS.EMAIL_OPT_OUT - 1] === true) {
  email = '(UNSUBSCRIBED)' + email;
}
```

#### Affected Files
- `Constants.gs` - Add EMAIL_OPT_OUT to MEMBER_COLS
- `MemberDirectory.gs` - Add column setup
- `ReportGeneration.gs` - Modify export functions
- `EmailIntegration.gs` - Filter opted-out members
- `DropdownConfig.gs` - Add checkbox validation
- `Code.gs` or Setup functions - Add conditional formatting

#### Testing Checklist
- [ ] Checkbox column appears in Member Directory
- [ ] Checking opt-out turns row light red
- [ ] Unchecking opt-out removes red background
- [ ] CSV export prefixes email with (UNSUBSCRIBED)
- [ ] Bulk email excludes opted-out members
- [ ] Warning shown when manually emailing opted-out member

---

## Completed Features

_(Move items here when implemented)_

---

## Notes

- Feature requests can be added to this file
- Prioritize based on user impact and implementation complexity
- Update status as work progresses: Planned -> In Progress -> Completed
