# SalaryCalc

SalaryCalc is a Google Apps Script-based solution for automating payroll calculations, attendance management, and payslip generation using Google Sheets. It is designed for organizations that want to streamline their monthly payroll process, track attendance, and generate professional payslips for employees.

## Features
- **Automated Payroll Calculation:** Calculates salaries, deductions, and net pay based on attendance and other parameters.
- **Attendance Management:** Applies formulas to attendance data, highlights missing punches, and updates attendance statuses.
- **Payslip Generation:** Generates PDF payslips for each employee and saves them to Google Drive.
- **Template Management:** Prepares the next month's template for loan and advance tracking.
- **Cell Navigation:** Easily navigate between highlighted cells for quick review and correction.

## Getting Started

### Prerequisites
- A Google account
- Access to Google Sheets
- Basic familiarity with Google Apps Script

### Setup Instructions
1. **Clone or Copy the Scripts:**
   - Copy the contents of each `.js` file into the Google Apps Script editor attached to your Google Sheet.
2. **Configure Script Properties:**
   - Set up any required script properties (e.g., `driveFolderId` for payslip storage) via the Apps Script dashboard.
3. **Sheet Structure:**
   - Ensure your Google Sheet has the following sheets: `Input Data`, `Final Sheet`, `Loan & Advance`, and `Logs`.
4. **Permissions:**
   - The script will request permissions to access your Google Drive and Sheets. Approve these when prompted.

## Usage
- **Payroll Menu:** After installation, a custom menu called `Payroll` will appear in your Google Sheet. Use it to generate payslips.
- **Attendance Processing:** Run the functions in `coreLogic.js` to apply formulas and update attendance data.
- **Highlight Missing Punches:** Use `findMissingPunches.js` to highlight and log missing attendance punches.
- **Navigate Highlights:** Use `cellNavigator.js` to move between highlighted cells for review.
- **Generate Next Month's Template:** Use `generateTemplate.js` to prepare the next month's loan and advance sheet.

## File Descriptions

- **coreLogic.js**
  - Contains the main logic for processing attendance data, applying formulas, updating statuses, and clearing or preparing data in the `Input Data` sheet.
  - Functions include: applying formulas for late/early marks, durations, overtime, updating attendance status, and more.

- **generatePayslips.js**
  - Handles the generation of PDF payslips for each employee based on the data in the `Final Sheet`.
  - Saves payslips to a specified Google Drive folder and formats salary details, deductions, and net pay.

- **generateTemplate.js**
  - Automates the creation of the next month's template in the `Loan & Advance` sheet.
  - Copies formulas, formatting, and headers, and updates the month value.

- **findMissingPunches.js**
  - Highlights cells in the `Input Data` sheet where attendance punches are missing after a 'P' (Present) status.
  - Logs the addresses of highlighted cells in the `Logs` sheet for easy review.

- **cellNavigator.js**
  - Provides functions to navigate to the next or previous highlighted cell in the `Input Data` sheet using the log in the `Logs` sheet.
  - Useful for quickly reviewing and correcting missing or problematic entries.

- **appsscript.json**
  - Google Apps Script project manifest file. Sets timezone, runtime, and logging options.

## License
This project is provided as-is for internal or educational use. Please review and adapt it to your organization's requirements before deploying in a production environment. 