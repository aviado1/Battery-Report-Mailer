# Load Outlook COM object
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# Specify the folder path to access the 'Battery-Report' folder
try {
    $reportsFolder = $namespace.Folders.Item("user@example.com").Folders.Item("Inbox").Folders.Item("Reports").Folders.Item("Battery-Report")
} catch {
    Write-Output "Failed to locate the 'Battery-Report' folder. Please verify the folder structure."
    exit
}

# Define the output folder path
$outputFolder = "C:\Temp"

# Ensure the output folder exists
if (-not (Test-Path -Path $outputFolder)) {
    New-Item -Path $outputFolder -ItemType Directory | Out-Null
}

# Iterate through each email in the 'Battery-Report' folder
foreach ($mailItem in $reportsFolder.Items) {
    if ($mailItem.Attachments.Count -gt 0) {
        foreach ($attachment in $mailItem.Attachments) {
            # Save the attachment to the output folder
            $filePath = Join-Path -Path $outputFolder -ChildPath $attachment.FileName
            $attachment.SaveAsFile($filePath)
            Write-Output "Saved attachment: $filePath"
        }
    }
}

# Analyze battery report HTML files and create an Excel report
$outputExcelPath = "C:\Temp\Excel_Report\Battery_Report_Analysis.xlsx"

# Load the required module for Excel manipulation
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

Import-Module ImportExcel

# Create an array to hold the data
$batteryData = @()

# Iterate through each battery report HTML file in the output folder
Get-ChildItem -Path $outputFolder -Filter "battery-report-*.html" | ForEach-Object {
    $htmlFile = $_.FullName
    $computerName = $_.BaseName -replace "battery-report-", ""

    # Read the HTML content
    $htmlContent = Get-Content -Path $htmlFile -Raw

    # Extract Design Capacity and Full Charge Capacity using regex
    $designCapacity = [regex]::Match($htmlContent, '(?i)DESIGN\s*CAPACITY.*?(\d{1,5},?\d{0,5})\s*mWh').Groups[1].Value -replace ",", ""
    $fullChargeCapacity = [regex]::Match($htmlContent, '(?i)FULL\s*CHARGE\s*CAPACITY.*?(\d{1,5},?\d{0,5})\s*mWh').Groups[1].Value -replace ",", ""

    # Convert to integers and calculate battery health percentage
    if ($designCapacity -and $fullChargeCapacity) {
        $designCapacity = [int]$designCapacity
        $fullChargeCapacity = [int]$fullChargeCapacity
        $batteryHealthPercentage = [math]::Round(($fullChargeCapacity / $designCapacity) * 100, 2)

        # Add the data to the array
        $batteryData += [PSCustomObject]@{
            ComputerName = $computerName
            DesignCapacity = $designCapacity
            FullChargeCapacity = $fullChargeCapacity
            BatteryHealth = $batteryHealthPercentage
        }
    }
}

# Sort the data by battery health
$batteryData = $batteryData | Sort-Object BatteryHealth

# Export the data to an Excel file with bold headers, centered alignment, and conditional formatting
$batteryData | Export-Excel -Path $outputExcelPath -AutoSize -WorksheetName "Battery Health Analysis" -BoldTopRow -TableName "BatteryHealthTable"

# Open the Excel file and apply additional formatting
$excel = Open-ExcelPackage -Path $outputExcelPath
$worksheet = $excel.Workbook.Worksheets["Battery Health Analysis"]

# Center align all columns
$worksheet.Cells.Style.HorizontalAlignment = "Center"

# Apply conditional formatting based on Battery Health
$condition = $worksheet.Cells[2, 4, $worksheet.Dimension.End.Row, 4].ConditionalFormatting.AddThreeColorScale()
$condition.LowValue.Color = "Red"  # Red for low health
$condition.MiddleValue.Color = "Yellow"  # Yellow for medium health
$condition.HighValue.Color = "Green"  # Green for high health

# Save and close the Excel file
Close-ExcelPackage -ExcelPackage $excel

Write-Output "Battery health analysis saved to: $outputExcelPath"
