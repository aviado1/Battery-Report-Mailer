# Version Number
$scriptVersion = "1.3"

# Define a log path within the C:\Temp directory
$logPath = "C:\Temp\battery-report-error.log"

# Log helper function
function Write-Log {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "${timestamp}: $message"
}

# Log the script version
Write-Log "Running Battery Report Mailer Version: $scriptVersion"

$maxRetries = 3
$retryCount = 0
$emailSent = $false

# Retrieve the computer name
$computerName = $env:COMPUTERNAME

# Define variables
$reportPath = "C:\Temp\battery-report-$computerName.html"
$recipient = "recipient@example.com"
$subject = "Automated Battery Report - $computerName (Version $scriptVersion)"
$alertThreshold = 80  # Set the battery health threshold percentage (Full Charge Capacity / Design Capacity)

# SMTP Server Configuration
$SMTPServer = "smtp.example.com"
$SMTPFrom = "BatteryReport@example.com"
$SMTPPort = 25

# Generate Battery Report Silently with computer name in the file
Start-Process -FilePath "powercfg" -ArgumentList "/batteryreport", "/output", "`"$reportPath`"" -NoNewWindow -Wait

# Check if the report was generated successfully
if (Test-Path $reportPath) {
    try {
        # Read the HTML content
        $htmlContent = Get-Content -Path $reportPath -Raw

        # Log the beginning of HTML parsing
        Write-Log "Parsing HTML content of the battery report."

        # Use improved regular expressions to find Design Capacity and Full Charge Capacity
        $designCapacityMatch = [regex]::Match($htmlContent, '(?i)DESIGN\s*CAPACITY.*?(\d{1,5},?\d{0,5})\s*mWh')
        $fullChargeCapacityMatch = [regex]::Match($htmlContent, '(?i)FULL\s*CHARGE\s*CAPACITY.*?(\d{1,5},?\d{0,5})\s*mWh')

        # Extract values if matched
        if ($designCapacityMatch.Success -and $fullChargeCapacityMatch.Success) {
            $designCapacity = $designCapacityMatch.Groups[1].Value -replace ",", ""
            $fullChargeCapacity = $fullChargeCapacityMatch.Groups[1].Value -replace ",", ""
        } else {
            Write-Log "Failed to find Design Capacity or Full Charge Capacity in the report."
            Write-Output "Failed to find Design Capacity or Full Charge Capacity in the report."
            throw "Failed to extract capacity values"
        }

        # Log and output the extracted values
        Write-Log "Extracted Design Capacity (raw): $designCapacity"
        Write-Log "Extracted Full Charge Capacity (raw): $fullChargeCapacity"
        Write-Output "Extracted Design Capacity (raw): $designCapacity"
        Write-Output "Extracted Full Charge Capacity (raw): $fullChargeCapacity"

        # Convert the captured values to integers
        try {
            $designCapacity = [int]$designCapacity
            $fullChargeCapacity = [int]$fullChargeCapacity
        }
        catch {
            Write-Log "Error converting capacity values to integers: $($_.Exception.Message)"
            Write-Output "Error converting capacity values to integers: $($_.Exception.Message)"
            throw "Integer conversion error"
        }

        # Log the numeric values after conversion
        Write-Log "Converted Design Capacity (int): $designCapacity"
        Write-Log "Converted Full Charge Capacity (int): $fullChargeCapacity"
        Write-Output "Converted Design Capacity (int): $designCapacity"
        Write-Output "Converted Full Charge Capacity (int): $fullChargeCapacity"

        # Check if Design Capacity is valid
        if ($designCapacity -eq 0) {
            Write-Output "Design Capacity is 0, cannot calculate battery health."
            Write-Log "Design Capacity is 0, cannot calculate battery health."
            throw "Design Capacity is zero"
        }

        # Calculate the percentage of battery health
        $batteryHealthPercentage = [math]::Round(($fullChargeCapacity / $designCapacity) * 100, 2)

        Write-Output "Calculated Battery Health Percentage: $batteryHealthPercentage%"
        Write-Log "Calculated Battery Health Percentage: $batteryHealthPercentage%"

        # Determine if the battery is OK or Not OK
        $batteryStatus = "OK"
        $statusColor = "green"
        if ($batteryHealthPercentage -lt $alertThreshold) {
            $batteryStatus = "NOT OK"
            $statusColor = "red"
        }

        # Create the HTML body for the email with color-coded status
        $body = @"
            <html>
            <body>
                <p>Please find the attached battery report for <b>$computerName</b>.</p>
                <p><b>Battery Health: <span style='color:$statusColor'>$batteryStatus</span></b> ($batteryHealthPercentage% of Design Capacity)</p>
                <p>This is an automated email from the battery report monitoring system (Version $scriptVersion).</p>
            </body>
            </html>
"@

        # Retry mechanism for sending email
        while ($retryCount -lt $maxRetries -and $emailSent -eq $false) {
            try {
                # Send the email with the battery report attached
                Send-MailMessage -From $SMTPFrom `
                                 -To $recipient `
                                 -Subject $subject `
                                 -BodyAsHtml $body `
                                 -SmtpServer $SMTPServer `
                                 -Port $SMTPPort `
                                 -Attachments $reportPath

                $emailSent = $true
                Write-Output "Email sent successfully."
                Write-Log "Email sent successfully."
            }
            catch {
                $retryCount++
                Write-Log "Attempt $retryCount failed: $($_.Exception.Message)"
                Start-Sleep -Seconds 5  # Wait 5 seconds before retrying
            }
        }

        if ($emailSent -eq $false) {
            Write-Log "All attempts to send email failed."
        }

    }
    catch {
        Write-Log "Error processing battery report: $($_.Exception.Message)"
    }
}
else {
    Write-Log "Battery report was not generated."
}
