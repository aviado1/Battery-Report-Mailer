
# Battery Report Scripts

**Author**: [aviado1](https://github.com/aviado1)

This repository contains two PowerShell scripts designed to generate and analyze battery reports from Windows machines. The first script, `BatteryReportMailer_v1.3.ps1`, generates battery reports and sends them via email. The second script, `BatteryReportOutlookToExcel.ps1`, retrieves the emailed battery reports from an Outlook folder, processes the data, and exports the analysis to an Excel file.

## How to Use

### Step 1: Generate Battery Report and Send via Email

1. **Suggestion**: Create a new folder in your Outlook named `Battery-Report` to store the reports you receive via email. You can also set up a rule in Outlook to automatically move emails with battery reports to this folder.

1. **Script**: `BatteryReportMailer_v1.3.ps1`
2. **Description**: This script generates a battery report for the local machine and sends it via email to the designated recipient.
3. **Instructions**:
   - Open PowerShell as an administrator.
   - Run the script to generate a battery report in HTML format and send it via email. You can run this script using any mechanism that supports remote PowerShell execution, such as Task Scheduler, a remote management tool, or a custom automation solution.
   - Make sure to adjust the SMTP server configuration and recipient details as needed.

### Step 2: Retrieve Battery Reports from Outlook and Export Analysis to Excel

1. **Script**: `BatteryReportOutlookToExcel.ps1`
2. **Description**: This script retrieves battery report attachments from a specified Outlook folder, analyzes the battery health, and exports the data to an Excel report.
3. **Instructions**:
   - Ensure you have Microsoft Outlook installed and configured.
   - Make sure the battery report emails are located in the `Battery-Report` folder within your Outlook mailbox.
   - Run the script to extract the battery reports, analyze their contents, and save the analysis to an Excel file.

## Prerequisites

- PowerShell 5.1 or higher.
- Outlook application installed and configured.
- [ImportExcel](https://www.powershellgallery.com/packages/ImportExcel) PowerShell module for Excel operations. The script will install it if it's not already available.

## Disclaimer

These scripts are provided as-is without any warranty. The author takes no responsibility for any issues that arise from using these scripts. Use at your own risk.
