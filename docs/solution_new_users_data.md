# Solution: New Users Data Display

This document explains the issues encountered when adding the "New Users" section and the robust solutions implemented to fix them.

## 1. Problem: "No Data Showing" on Screen

Even when the Excel file had data, the web view often showed an empty "New Users" table. This was caused by several factors:

### A. Non-Standard Sheet Names
**Issue:** The code looked for exactly "New Users", but some files used "new users" or "New users".
**Solution:** Implemented case-insensitive sheet lookup in `report_parser.py`.

```python
# Before
if "New Users" in all_sheets:

# After
new_users_sheet_name = next((name for name in all_sheets.keys() if name.lower() == "new users"), None)
if new_users_sheet_name:
```

### B. Shifted Header Rows (Junk Data at Top)
**Issue:** Many Excel exports have 2-3 empty rows or titles at the very top. Pandas would read these as headers, naming columns `Unnamed: 0`, `Unnamed: 1`, etc.
**Solution:** Added a **Header Scanner** that searches the first 10 rows for keywords like "Ticket", "Date", "User", or "Email" to find the real header row.

### C. Column Name Variations
**Issue:** One file might use "Name", another "User Name", and another "UserName". The UI would fail to find the data if the keys didn't match.
**Solution:** Implemented a **Column Mapping Dictionary** to normalize common variations into standard keys.

## 2. Problem: Frontend Rendering Failures

### A. Missing Sort State
**Issue:** The JavaScript `handleSort` function crashed because `sortState.newUsersTable` was undefined. In JavaScript, a single error in a click handler can stop all subsequent rendering.
**Solution:** Initialized the sort state for the new table.

```javascript
let sortState = {
    leftTable: { col: null, asc: true },
    rightTable: { col: null, asc: true },
    newUsersTable: { col: null, asc: true } // Added this
};
```

### B. Rigid Column Expectations
**Issue:** The frontend was hardcoded to look for exactly 5 specific columns. If the backend sanitized a name or the Excel had an extra column, the row values would be blank.
**Solution:** Implemented **Dynamic Column Detection**. The UI now reads the actual keys from the JSON data and builds the table headers on the fly.

## 3. Best Practices for Future Sections

If a new section (e.g., "Monthly Totals") is added in the future, follow these steps to avoid similar issues:

1.  **Backend (Service Layer):** Use a keyword-based scanner to find headers rather than assuming they are at row 0.
2.  **DTO/JSON Layer:** Ensure all `NaN` or `NaT` values are converted to `None` (null) so the JSON is valid and easy for JavaScript to handle.
3.  **Frontend (UI Layer):** 
    *   Always initialize `sortState` for new tables.
    *   Generate `<th>` elements dynamically based on the object keys.
    *   Use a fallback lookup for values (e.g., `row[col] || row[col.replace('.', '')]`).
4.  **Integration (PDF Layer):** Use the same dynamic column list for the PDF table to ensure the Web and PDF views match perfectly.
