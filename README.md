# UBL L&D Executive Dashboard
## Setup & Deployment Guide

---

### Files in This Package

| File | Purpose |
|------|---------|
| `dashboard.html` | Main executive dashboard — share this with C-level |
| `admin.html` | Admin panel — for L&D team to upload Excel files |
| `Code.gs` | Google Apps Script backend (optional, for cloud storage) |

---

### Quick Start (GitHub Pages — Browser Storage)

This is the simplest setup. Data is stored in the browser (localStorage).

1. Create a new GitHub repository (e.g. `ubl-ld-dashboard`)
2. Upload both `dashboard.html` and `admin.html`
3. Go to **Settings → Pages → Source: main branch → / (root)**
4. Your dashboard will be live at:
   `https://yourusername.github.io/ubl-ld-dashboard/dashboard.html`
5. Admin panel:
   `https://yourusername.github.io/ubl-ld-dashboard/admin.html`

**Default admin credentials:**
- Username: `admin`
- Password: `ubl2025`
- *(Change these immediately from the Admin Panel → Settings)*

---

### How to Upload Data

1. Open the Admin Panel (`admin.html`)
2. Log in with admin credentials
3. Enter the **Month Label** (e.g. `March 2025`)
4. Drag & drop your Excel file OR click Browse
5. Review the data preview and column mapping
6. Click **Save to Repository & Update Dashboard**

The dashboard will instantly reflect the new data.

---

### Excel File Format

Your Excel file should contain these columns (column names are flexible — the system will auto-map):

| Column | Expected Values |
|--------|----------------|
| Employee No. | 6-digit ID |
| Employee Name | Full name |
| Functional Title | Job title |
| Grade | OG-4 to SEVP-II |
| Group | Department/group |
| POP Code | 4-digit branch code |
| Cluster | South / Central / North |
| Attendance Status | Enrolled / Attended (or similar) |
| Modules | Module name |
| Course Title | Program title |
| Delivery Mode | Virtual / In-House |
| Training Type | Online / Classroom |
| Training Status | Planned / Unplanned |
| From | Start date |
| To | End date |
| Duration | Number of days |
| Working Hours | Hours per day (usually 8) |
| Category | Functional & Technical / UBL Specific |

---

### Google Sheets Backend (Optional — for team-wide access)

If multiple team members need to upload data from different devices, use the Google Apps Script backend:

1. Open **Google Sheets** → create a new blank spreadsheet
2. Copy the **Spreadsheet ID** from the URL
3. Go to **Extensions → Apps Script**
4. Paste the contents of `Code.gs` into the editor
5. Replace `SPREADSHEET_ID = ''` with your sheet's ID
6. Click **Deploy → New Deployment → Web App**
   - Execute as: **Me**
   - Who has access: **Anyone within your organization**
7. Copy the Web App URL
8. In `admin.html`, find the line:
   ```javascript
   const APPS_SCRIPT_URL = '';
   ```
   And paste your URL there.

---

### Changing Admin Password

1. Log in to Admin Panel
2. Scroll to **Admin Settings** at the bottom
3. Enter new username and/or password
4. Click **Save Credentials**

---

### Printing / PDF Export

From the main dashboard:
- Click **Print** for browser print dialog
- Click **Export PDF** → choose "Save as PDF" in the print dialog

For best PDF output: set margins to Minimum, enable Background graphics.

---

### Support

For technical issues or feature requests, contact the L&D Team.
UBL Learning & Development Training Center — 2025
