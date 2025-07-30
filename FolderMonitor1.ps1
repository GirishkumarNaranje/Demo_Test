Clear-Host

# Configuration
$folderToMonitor = "D:\WebSites\itclearningU\Site_DemoNET\JCAScorm"
$script1 = "D:\PROD\jobs\job-fix-logout.ps1"
$script2 = "D:\PROD\jobs\job-fix-youcanclosethispopup.ps1"
$logFolder = "D:\PROD\jobs\FolderMonitorLogs"
$stateFile = Join-Path $logFolder "digit_folders_state.json"

# Email settings - Office 365 Example
#$emailFrom = "Girishkumar.Naranje@Harbingergroup.com"
#$emailTo = "Girishkumar.Naranje@Harbingergroup.com"
#$emailSubject = "ITC Training - Folder Monitor Script execution for course changes/added Log"
#$smtpServer = "smtp.office365.com"
#$smtpPort = 587
#$smtpUser = "Girishkumar.Naranje@Harbingergroup.com"
#$smtpPass = "lhptmmrlkgpznmhn"  # Use an App Password if MFA is enabled!


# Email settings
$emailFrom = "support@itclearning.com"
$emailTo = @("girishkumar.naranje@harbingergroup.com", "naranje.girish@gmail.com")
$emailSubject = "Course Folder Update Detected - Automation Script Execution Log"
$smtpServer = "email-smtp.us-east-1.amazonaws.com"
$smtpPort = 587 
$smtpUser = "AKIAJXNJZXWNBOKOYJXQ"
$smtpPass = "BLqrbnuujpgico0dh5QoFt6j4vFQ5ZgD04krxSF+8YdH" 

if (-Not (Test-Path $logFolder)) { New-Item -Path $logFolder -ItemType Directory | Out-Null }

function Write-Log {
    param([string]$msg)
    $logFile = Join-Path $logFolder ("MonitorFolderChanges_{0}.log" -f (Get-Date -Format "yyyy-MM-dd"))
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "$timestamp - $msg"
}

function Get-DigitFoldersState {
    $dirs = Get-ChildItem -Path $folderToMonitor -Directory | Where-Object { $_.Name -match '^\d+$' }
    $state = @{}
    foreach ($dir in $dirs) {
        $times = @($dir.CreationTime, $dir.LastWriteTime)
        $fileTimes = Get-ChildItem -Path $dir.FullName -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { -not $_.PSIsContainer } | ForEach-Object { @($_.CreationTime, $_.LastWriteTime) }
        $maxTime = ($times + $fileTimes | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1)
        $state[$dir.FullName] = $maxTime
    }
    return $state
}

function Load-PreviousState {
    if (Test-Path $stateFile) {
        try {
            $raw = Get-Content $stateFile | ConvertFrom-Json
            $converted = @{}
            foreach ($key in $raw.PSObject.Properties.Name) {
                $converted[$key] = [DateTime]::Parse($raw.$key)
            }
            return $converted
        } catch {
            Write-Log "WARN: Failed to read previous state, using empty state."
            return @{}
        }
    } else {
        return @{}
    }
}

function Save-State ($state) {
    $serializable = @{}
    foreach ($k in $state.Keys) {
        $serializable[$k] = $state[$k].ToString("o")
    }
    $serializable | ConvertTo-Json | Set-Content $stateFile
}

# MAIN LOGIC
Write-Log "Script started - FolderMonitor.ps1"

if (-Not (Test-Path $folderToMonitor)) { Write-Log "ERROR: Folder not found: $folderToMonitor"; exit 1 }
if (-Not (Test-Path $script1)) { Write-Log "ERROR: Script1 not found: $script1"; exit 1 }
if (-Not (Test-Path $script2)) { Write-Log "ERROR: Script2 not found: $script2"; exit 1 }

$prevState = Load-PreviousState
$currState = Get-DigitFoldersState

$counter = 0
foreach ($dir in $currState.Keys) {
    if (-not $prevState.ContainsKey($dir)) {
        Write-Log "New folder: $dir"
        $counter++
    } elseif ($currState[$dir] -gt $prevState[$dir]) {
        Write-Log "Change in: $dir"
        $counter++
    }
}

$allScriptsSucceeded = $true

if ($counter -gt 0) {
    Write-Log "$counter change(s) detected. Running scripts."
    try {
        Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$script1`"" -WindowStyle Hidden -Wait
        Write-Log "Script1 - job-fix-logout.ps1 completed."
    } catch {
        Write-Log "ERROR: Failed Script1: $_"
        $allScriptsSucceeded = $false
    }
    try {
        Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$script2`"" -WindowStyle Hidden -Wait
        Write-Log "Script2 - job-fix-youcanclosethispopup.ps1 completed."
    } catch {
        Write-Log "ERROR: Failed Script2: $_"
        $allScriptsSucceeded = $false
    }
} else {
    Write-Log "No changes detected."
}

Save-State $currState
Write-Log "Script completed - FolderMonitor.ps1"

# --- EMAIL LOGIC ---
if ($allScriptsSucceeded -and $counter -gt 0) {
    $logFile = Join-Path $logFolder ("MonitorFolderChanges_{0}.log" -f (Get-Date -Format "yyyy-MM-dd"))
    $logContent = Get-Content $logFile -Raw

    # Custom message and HTML formatting
    $customMessage = "Hello,<br><br>An automated system has detected changes in the course folder structure. Below are the execution logs of the automation script.<br><br>
	Log Format: ExecutionDate ExecutionTime - Action.<br><br>
	Change in: Indicates modifications in existing course folders.<br>
	New folder: Indicates creation of new course folders.<br>
	<br><b>Logs:</b>"
    $emailBody = $customMessage + "<pre>" + $logContent + "</pre>"
	$emailBody += '<br><br><small><i>This is an automated notification. Please do not reply to this email.</i></small>'

    # Use .NET MailMessage for full HTML support
    $mail = New-Object System.Net.Mail.MailMessage
    $mail.From = $emailFrom
    #$mail.To.Add($emailTo)
	foreach ($recipient in $emailTo) {
		$mail.To.Add($recipient)
	}
    $mail.Subject = $emailSubject
    $mail.Body = $emailBody
    $mail.IsBodyHtml = $true

    $smtp = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort)
    $smtp.EnableSsl = $true
    $smtp.Credentials = New-Object System.Net.NetworkCredential($smtpUser, $smtpPass)

    try {
        $smtp.Send($mail)
        Write-Log "Log email sent to $emailTo."
    } catch {
        Write-Log "ERROR: Failed to send log email: $_"
    }
}