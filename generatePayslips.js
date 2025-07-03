function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Payroll')
      .addItem('Generate Payslip', 'generatePayslips')
      .addToUi();
}

function generatePayslips() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Final Sheet");
    const driveFolderId = PropertiesService.getScriptProperties().getProperty('driveFolderId');
    const parentFolder = DriveApp.getFolderById(driveFolderId);
    
    // Get Salary Month from E3
    const salaryMonth = formatMonthYear(sheet.getRange("E1").getValue());
    if (!salaryMonth) {
        Logger.log("Salary month not found in E1.");
        return;
    }

    // Create or get folder for the month
    const monthFolder = getOrCreateFolder(parentFolder, salaryMonth);
    
    // Get total working days once
    const totalWorkingDays = sheet.getRange("Q1").getValue();
    
    // Get all data at once for efficiency
    const data = sheet.getDataRange().getValues();
    
    // Start from row 3 (index 2) which contains the first employee
    for (let i = 2; i < data.length; i++) {
        const row = data[i];
        const empCode = row[3]; // Column D
        const empName = row[4]; // Column E

        if (!empCode || !empName) continue;

        const actualPayableDays = row[9]; // Column J
        const lossOfPayDays = totalWorkingDays - actualPayableDays;
        const netSalary = row[23]; // Column X - Net Payable

        const payslipData = {
            "Company": row[1],
            "Employee Code": empCode,
            "Employee Name": empName,
            "Basic Salary": formatCurrency(row[6]), // Column G
            "Present Days": formatNumber(actualPayableDays, 0),
            "Total Working Days": formatNumber(totalWorkingDays, 0),
            "Loss Of Pay Days": formatNumber(lossOfPayDays, 0),
            "Paid Present Days": formatCurrency(row[10], 0), // Column K
            "OT Hours": formatNumber(row[11]), // Column L
            "Late Mark Hours": formatNumber(row[12]), // Column M
            "Early Going Hours": formatNumber(row[13]), // Column N
            "Actual Worked Hours": formatNumber(row[14]), // Column O
            "Paid Worked Hours": formatNumber(row[15]), // Column P
            "Gross Wages": formatCurrency(row[16]), // Column Q
            "Conveyance Allowance": formatCurrency(row[17]), // Column R
            "Loan Deduction": formatCurrency(row[20]), // Column U
            "Food Expense": formatCurrency(row[18]), // Column S
            "Net Amount": formatCurrency(row[21]), // Column V
            "Rounded Off": formatCurrency(row[22]), // Column W
            "Net Salary": formatCurrency(netSalary),
            "Net Salary Words": numberToWords(Math.round(netSalary)) // Convert to words
        };

        const pdfBlob = createPayslipPDF(empName, salaryMonth, payslipData);
        if (pdfBlob) {
            monthFolder.createFile(pdfBlob);
        }
    }

    Logger.log("Payslips generated successfully.");
}

// Helper function to get or create a folder
function getOrCreateFolder(parentFolder, folderName) {
    const folders = parentFolder.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

// Function to create the payslip PDF
function createPayslipPDF(empName, salaryMonth, data) {
    const htmlContent = `
    <html>
    <head>
        <style>
            body { font-family: Arial, sans-serif; font-size: 12px; margin: 0; padding: 0; }
            .container { width: 210mm; height: 297mm; padding: 20mm; box-sizing: border-box; }
            .header { text-align: left; font-size: 24px; font-weight: bold; margin-bottom: 10px; }
            .sub-header { text-align: left; font-size: 20px; margin-bottom: 20px; }
            .salary-details { width: 100%; margin-top: 20px; margin-bottom: 20px; }
            .salary-details-table { width: 100%; border-collapse: collapse; }
            .salary-details-table td { padding: 2px; text-align: left; }
            .salary-details-table .right-value { text-align: right; }
            .working-hours-table { width: 100%; border-collapse: collapse; }
            .working-hours-table td { padding: 2px; text-align: left; }
            .working-hours-table .right-value { text-align: right; }
            
            hr { border: none; border-top: 1px solid #000; margin: 15px 0; }
            .salary-section-header { 
                font-size: 16px; 
                font-weight: bold; 
                margin: 15px 0 10px 0; 
                background-color: #f2f2f2;
                padding: 5px;
            }
            
            .earnings-deductions-container {
                display: flex;
                width: 100%;
                margin-top: 20px;
            }
            
            .earnings-section {
                width: 48%;
            }
            
            .deductions-section {
                width: 48%;
                margin-left: 4%;
            }
            
            .earnings-deductions-table {
                width: 100%;
                border-collapse: collapse;
            }
            
            .earnings-deductions-table td {
                padding: 5px 0;
                vertical-align: top;
            }
            
            .item-name {
                text-align: left;
                font-weight: normal;
            }
            
            .item-value {
                text-align: right;
                font-weight: normal;
            }
            
            .total-row td {
                font-weight: bold;
                border-top: 1px solid #000;
                padding-top: 8px;
            }
            
            .net-salary-container {
                background-color: #f2f2f2;
                padding: 10px;
                margin-top: 20px;
            }
            
            .net-salary-table {
                width: 100%;
            }
            
            .net-salary-table td {
                padding: 5px 0;
            }
            
            .footer { text-align: center; font-size: 10px; margin-top: 20px; font-style: italic; }
            .divider-line {
                width: 1px;
                background-color: #000;
                margin: 0 10px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">PAYSLIP ${salaryMonth}</div>
            <div class="sub-header">${data["Company"]}</div>

            <table style="width: 100%; margin-bottom: 20px;">
                <tr>
                    <td style="width: 25%; font-weight: bold;">Employee Code</td>
                    <td style="width: 25%;">${data["Employee Code"]}</td>
                    <td style="width: 25%; font-weight: bold;">Employee Name</td>
                    <td style="width: 25%;">${data["Employee Name"]}</td>
                </tr>
            </table>

            <hr>

            <div class="salary-section-header">SALARY DETAILS</div>
            
            <table class="salary-details-table">
                <tr>
                    <th style="width: 25%;">Actual Payable Days</th>
                    <th style="width: 25%;">Loss Of Pay Days</th>
                    <th style="width: 25%;">Days Payable</th>
                    <th style="width: 25%;">Total Working Days</th>
                </tr>
                <tr>
                    <td>${data["Present Days"]}</td> 
                    <td>${data["Loss Of Pay Days"]}</td>
                    <td>${data["Present Days"]}</td>
                    <td>${data["Total Working Days"]}</td>
                </tr>
            </table>

            <hr>

            <div class="salary-section-header">WORKING HOURS</div>
            
            <table class="working-hours-table">
                <tr>
                    <th style="width: 25%;">OT (Hours)</th>
                    <th style="width: 25%;">Late Mark (Hours)</th>
                    <th style="width: 25%;">Early Going (Hours)</th>
                    <th style="width: 25%;">Actual Worked Hours</th>
                </tr>
                <tr>
                    <td>${data["OT Hours"]}</td> 
                    <td>${data["Late Mark Hours"]}</td>
                    <td>${data["Early Going Hours"]}</td>
                    <td>${data["Actual Worked Hours"]}</td>
                </tr>
            </table>

            <hr>

            <div class="earnings-deductions-container">
                <div class="earnings-section">
                    <div class="salary-section-header">EARNINGS</div>
                    <table class="earnings-deductions-table">
                        <tr>
                            <td class="item-name">Paid Present Days</td>
                            <td class="item-value">${data["Paid Present Days"]}</td>
                        </tr>
                        <tr>
                            <td class="item-name">Paid Worked Hours</td>
                            <td class="item-value">${data["Paid Worked Hours"]}</td>
                        </tr>
                        <tr>
                            <td class="item-name">Conveyance Allowance</td>
                            <td class="item-value">${data["Conveyance Allowance"]}</td>
                        </tr>
                        <tr>
                            <td class="item-name">Food Expense</td>
                            <td class="item-value">${data["Food Expense"]}</td>
                        </tr>
                        <tr class="total-row">
                            <td class="item-name">Total Earnings (A)</td>
                            <td class="item-value">${data["Gross Wages"]}</td>
                        </tr>
                    </table>
                </div>
                
                <div class="deductions-section">
                    <div class="salary-section-header">DEDUCTIONS</div>
                    <table class="earnings-deductions-table">
                        <tr>
                            <td class="item-name">Loan Deduction</td>
                            <td class="item-value">${data["Loan Deduction"]}</td>
                        </tr>
                        <tr class="total-row">
                            <td class="item-name">Total Deductions (C)</td>
                            <td class="item-value">${data["Loan Deduction"]}</td>
                        </tr>
                    </table>
                </div>
            </div>

            <div class="net-salary-container">
                <table class="net-salary-table">
                    <tr>
                        <td style="width: 50%; font-weight: bold;">Net Amount ( A - C )</td>
                        <td style="width: 50%; text-align: right; font-weight: bold;">${data["Net Amount"]}</td>
                    </tr>
                    <tr>
                        <td style="width: 50%; font-weight: bold;">Rounded Off</td>
                        <td style="width: 50%; text-align: right; font-weight: bold;">${data["Rounded Off"]}</td>
                    </tr>
                    <tr>
                        <td style="width: 50%; font-weight: bold;">Net Payable Salary</td>
                        <td style="width: 50%; text-align: right; font-weight: bold;">${data["Net Salary"]}</td>
                    </tr>
                    <tr>
                        <td style="font-weight: bold;">Net Salary in words</td>
                        <td style="text-align: right;">${data["Net Salary Words"]}</td>
                    </tr>
                </table>
            </div>
            
            <div class="footer">** Note : All amounts displayed in this payslip are in INR **</div>
            <div class="footer">** This is a computer-generated payslip and does not require a signature. **</div>
        </div>
    </body>
    </html>`;

    return HtmlService.createHtmlOutput(htmlContent).getBlob().getAs('application/pdf').setName(`${empName}.pdf`);
}

function formatMonthYear(date) {
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "MMMM-yyyy");
}

function formatCurrency(value) {
    if (value === undefined || value === null || value === "" || isNaN(value)) {
        return "0.00";
    }
    return parseFloat(value).toFixed(2);
}

function formatNumber(value, decimalPlaces = 2) {
    if (value === undefined || value === null || value === "" || isNaN(value)) {
        return decimalPlaces === 0 ? "0" : "0.00";
    }
    return parseFloat(value).toFixed(decimalPlaces);
}

// Improved function to convert number to words for Indian Rupees
function numberToWords(num) {
    if (isNaN(num) || num === null || num === undefined) return "Zero Rupees only";
    
    // Handle decimal part
    num = parseFloat(num).toFixed(2);
    const parts = num.toString().split('.');
    const wholePart = parseInt(parts[0]);
    const decimalPart = parseInt(parts[1]);
    
    if (wholePart === 0 && decimalPart === 0) return "Zero Rupees only";
    
    const a = ['', 'One ', 'Two ', 'Three ', 'Four ', 'Five ', 'Six ', 'Seven ', 'Eight ', 'Nine ', 'Ten ', 
               'Eleven ', 'Twelve ', 'Thirteen ', 'Fourteen ', 'Fifteen ', 'Sixteen ', 'Seventeen ', 'Eighteen ', 'Nineteen '];
    const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
    
    function convertLessThanThousand(n) {
        if (n === 0) return '';
        if (n < 20) return a[n];
        if (n < 100) {
            return b[Math.floor(n / 10)] + (n % 10 !== 0 ? ' ' + a[n % 10] : '');
        }
        return a[Math.floor(n / 100)] + 'Hundred ' + (n % 100 !== 0 ? 'and ' + convertLessThanThousand(n % 100) : '');
    }
    
    // Indian number system: crore, lakh, thousand, hundred
    function convertIndian(n) {
        if (n === 0) return '';
        let str = '';
        
        // Work with crores (10,000,000)
        if (n >= 10000000) {
            str += convertLessThanThousand(Math.floor(n / 10000000)) + 'Crore ';
            n %= 10000000;
        }
        
        // Work with lakhs (100,000)
        if (n >= 100000) {
            str += convertLessThanThousand(Math.floor(n / 100000)) + 'Lakh ';
            n %= 100000;
        }
        
        // Work with thousands
        if (n >= 1000) {
            str += convertLessThanThousand(Math.floor(n / 1000)) + 'Thousand ';
            n %= 1000;
        }
        
        // Work with hundreds and remaining
        if (n > 0) {
            if (str !== '') str += 'and ';
            str += convertLessThanThousand(n);
        }
        
        return str;
    }
    
    let result = convertIndian(wholePart);
    
    // Add rupees
    result += 'Rupees ';
    
    // Add paise if any
    if (decimalPart > 0) {
        result += 'and ' + convertLessThanThousand(decimalPart) + 'Paise';
    }
    
    return result.trim() + ' only';
}