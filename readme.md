# Service Team Project Roster

A portable resource-planning tool with the **26 employees from your supplied workbook**. Names, job titles, departments, joining dates, employee types and visa issuance companies are preserved. Role is the initial skills value; add actual certifications and skills in Employee master.

No customer, project, assignment or leave details were present in the Excel. Those lists start empty. No fictional operational records are mixed with your employee list.

## Sabse pehle kaise use karein

1. Chaaron files ek folder mein rakhein: `index.html`, `data.json`, `readme.md`, `convert.py`.
2. `index.html` double-click karein. Aapke 26 employees already nazar aayenge.
3. **Settings & data** mein apne actual working days set karein. Default Monday–Friday hai.
4. **Customers → + Customer** se customer add karein.
5. **Projects → + Project** se customer select karke project, dates, location aur manager add karein.
6. **Leave & absence** mein known leave record karein; Approved leave availability block karti hai.
7. **Find available people** mein next project ki start/end dates, role/skill search aur required allocation daalein.
8. Suitable employee ke saamne **Assign** karein; project, task, dates aur allocation save karein.
9. **Employee roster** mein current customer/project, utilization, booking end aur next full-capacity date dekhein. **Capacity timeline** mein date-wise bookings aur leave dekhein.
10. Roz ka kaam update karne ke baad **Backup JSON** download karein.

## Files

| File | Purpose |
| --- | --- |
| `index.html` | Entire offline application: styles, JavaScript and embedded initial employee data; no external libraries or fonts required |
| `data.json` | Initial structured data and source metadata; also the format for backups and restore |
| `readme.md` | Setup, working rules and maintenance instructions |
| `convert.py` | Imports a new Excel employee list, merges with existing roster backups, or runs a local server |

## Screens and features

- **Resource overview:** active/joined headcount, booked capacity, employees with spare capacity, approved leave and current conflicts.
- **Employee roster:** search by name, role, skills, company or ID; filter by customer, role and status; select any as-of date; export filtered results as Excel-compatible CSV or print.
- **Capacity timeline:** 7, 14 or 30 days; previous/next periods; leave, free capacity, bookings and conflict indicators; hover for details and click to create an assignment.
- **Find available people:** checks every working day in a date range, including future bookings and approved leave. Supports 1–100% required allocation and a maximum 366-day planning window.
- **Assignments:** create, edit or delete project/task bookings; site, notes, status and inclusive dates; Excel CSV export.
- **Projects:** customer, project/PO reference, dates, location, manager, priority and status.
- **Customers:** customer and contact details, notes and linked project count.
- **Employee master:** add/edit employees, skills, joining date, department, visa company, employment type, active/inactive state and capacity. Linked records cannot be deleted accidentally.
- **Leave & absence:** Annual, Sick, Emergency, Unpaid, Training or Other absence; Pending/Approved/Cancelled status; return date and next full-capacity date.
- **Settings & data:** configurable working week and title; full JSON backup/restore; loading a regenerated data file.

## Capacity rules — read before planning

**Dates are inclusive.** A booking ending on Friday keeps the employee booked on Friday. The next available date respects the configured working calendar. The initial as-of date uses Dubai time.

**Allocation means percentage of one full-time employee.** At 100% capacity, two concurrent assignments of 60% and 40% are valid; 60% plus 50% is blocked. An employee configured at 50% capacity cannot accept a 100% booking. The availability planner reports the minimum spare capacity across the entire requested working-date range.

**Approved absence blocks full capacity.** Pending/Cancelled absence does not block. A new assignment overlapping approved absence is blocked. If genuine leave arises during an existing booking, the user can explicitly acknowledge the overlap when recording approved leave. The timeline and roster then show a conflict until the booking is reassigned, reduced in date range or cancelled. Duplicate overlapping approved leave is rejected when entered through forms.

**Planned, Confirmed and Completed assignment records consume capacity on their recorded dates.** Cancelled assignments do not. Completed preserves historical utilization. If work ends early, shorten the end date and then mark Completed. Future dates on a Completed booking still reserve capacity until corrected.

**Project status is informational.** Marking a project Completed, Cancelled or On hold does not silently erase its team bookings. Update the individual assignments as well. New assignments are offered only for open projects; existing records remain editable. Assignment dates must fit within project dates and cannot begin before an employee's joining date.

**Next free date** is the first working day on or after the selected date with the employee's full configured capacity available. The search horizon is three years; a dash means no result or an inactive employee. It does not guarantee availability throughout a later project. Always use the date-range planner for that decision.

**No bookings is not operational confirmation of availability.** It only means no commitment has been recorded. Import actual assignments and leave before making deployment decisions. Skill search does not verify qualifications or certification validity.

**Overview booked capacity** = total recorded assignment allocation on the selected working day ÷ capacity of active/joined employees who are not on approved leave. Conflicts can push the ratio above 100%. Non-working days show no utilization; capacity percent is not hours worked or payroll attendance.

## Saving, backups and sharing

Changes are automatically saved in **localStorage in the browser being used**. Export JSON regularly. If browser storage is unavailable, the app warns you; export before closing.

- This version has **no central database, login or automatic multi-user synchronization**.
- It does **not** write into `data.json` on your hard drive automatically.
- Keep one designated coordinator/master backup to avoid multiple people editing conflicting copies.
- Clearing browser data, switching browser/profile, changing the hosting origin or sometimes moving the HTML can make local records unavailable. Restore the latest JSON backup.
- **Import JSON** validates the schema and references, then asks before replacing the entire current workspace. It does not merge. Export the current roster first.
- Imported files can contain existing capacity or leave conflicts. The app displays these so the coordinator can resolve them; assignment forms prevent new overbookings.
- CSV exports open in Excel. These are reports, not restorable backups. Complete JSON retains IDs, links, settings and every record.
- File contents contain employee details. Store and distribute your backups only to authorized people.

### Direct-open vs local server

Double-clicking `index.html` uses the embedded initial employee data when there is no saved browser roster. Use **Import JSON** to load the accompanying data file or a newer backup.

For local HTTP operation with Python installed:

```sh
python convert.py --serve
```

Open `http://localhost:8000`. The first load fetches adjacent `data.json`; after the first edit, saved browser data takes precedence. To switch to a newer adjacent file, use **Settings & data → Load adjacent data.json** and confirm replacement.

Use `python convert.py --serve --port 8080` if port 8000 is occupied. The server binds only to localhost and stops with Ctrl+C. It serves files; it does not save edits to disk. Hosting these static files on cPanel or another web server follows the same browser-only persistence rules; do not expose employee data on an unrestricted public website.

## Convert a future Excel employee list

Python 3.9+ and `openpyxl` are needed **only for Excel conversion**. The HTML itself needs no Python or installation.

```sh
python -m pip install openpyxl
python convert.py "Existing Service Team Members & Newly hired.xlsx" -o data.json
```

Supported headers: `EMP. NAME` / `Employee Name` / `Name`, `JOB TITLE` / `Designation` / `Role`, `DEPARTMENT`, `JOINING DATE`, `TYPE`, `VISA ISSUANCE COMPANY` / `Company`, optional `Skills` / `Skillset`. It detects the header after title rows and reads all sheets with matching headers. Store joining dates as Excel dates or ISO `YYYY-MM-DD`. Duplicate names are rejected to avoid linking the wrong person.

**Important:** ordinary conversion produces a fresh employee master with no operational bookings. To preserve current roster records, download the current JSON backup and merge:

```sh
python convert.py "updated-employees.xlsx" --merge "roster-backup-2026-10-06.json" -o data.json
```

The converter matches normalized employee names, preserves existing IDs/assignments/leave/customers/projects and adds new employee IDs. Employees absent from the new Excel are retained; mark them inactive explicitly in the UI if required. A renamed employee is treated as a new employee, so correct/match the name before merging. Two employees with the same normalized name require manual ID-based management in the app rather than this name-based converter.

Conversion updates Excel-supplied employee fields; manually added skills survive when the workbook has no Skills column. Existing capacity, active state and notes are preserved. Import the result into the app and review conflicts and field changes. Regenerating `data.json` does not change the embedded fallback in `index.html`; Import JSON is the reliable way to load updated employees when double-clicking the file.

## Data schema

Schema version: `1`.

- `settings`: title and `workDays` numbers (`0` Sunday through `6` Saturday).
- `employees`: unique ID, name, role, skills, department, joiningDate, type, company, active, capacity, notes.
- `clients`: ID, name, contact, email, phone, notes.
- `projects`: ID, clientId, name, code, start, end, location, manager, priority, status, notes.
- `assignments`: ID, employeeId, projectId, task, start, end, allocation, status, location, notes.
- `leaves`: ID, employeeId, type, start, end, status, notes.
- `source`: source filename and conversion metadata.

IDs connect all records and must remain stable. Do not edit IDs manually in Excel or JSON. This is a daily planning roster, not a timesheet, payroll system or certification register.
