###################################################################
# Script to audit a server
# Requires Powershell 5
#
###################################################################

Write-Host "Starting server check script"

# Check if Module is installed, register with Microsoft Update, and scan for missing hotfixes
Write-Host "Checking to see if PSWindowsUpdate is already installed..."
if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
    Write-Host "    Module PSWindowsUpdate already installed."
} else {
    Write-Host "    Module PSWindowsUpdate is not installed. Installing now."
    Install-Module PSWindowsUpdate
    Add-WUServiceManager -ServiceID 7971f918-a847-4430-9279-4a52d1efe18d
}
Write-Host "`n"
Write-Host "Checking for missing updates..."
Get-WUList –MicrosoftUpdate
Write-Host "`n"

# Check if there are any existing Windows Backup backup sets
Write-Host "`n"
Write-Host "Checking for Windows Backup sets..."
Write-Host "`n"
get-wbbackupset

# Audit Security log for event 4625 (failed login attempt)
Write-Host "`n"
Write-Host "Checking for failed logins..."
Write-Host "`n"
Get-EventLog -LogName Security -InstanceId 4625 |
    Select-Object -Property TimeGenerated,
    @{N="AccountName";E={$_.Message.Split("`n")[12].Replace("Account Name:",$Null).Trim()}},
    @{N="Domain";E={$_.Message.Split("`n")[13].Replace("Account Domain:",$Null).Trim()}},
    @{N="Source";E={$_.Message.Split("`n")[26].Replace("Source Network Address:   ",$Null).Trim()}}







#$yesterday = (get-date) - (New-TimeSpan -day 1)


#"Application","Security","System" | ForEach-Object { 
#    Get-Eventlog -Newest 10 -LogName $_ 
#}

#Get-WinEvent application | Where-Object {$_.LevelDisplayName -eq "Error" -or $_.LevelDisplayName -eq "Warning"}




# List all connected console and RDP sessions
Write-Host "`n"
Write-Host "Listing connected user sessions..."
Write-Host "`n"
qwinsta


# List attached storage, size, and free space
Write-Host "`n"
Write-Host "Listing atatched volumes..."
Write-Host "`n"
Get-Volume | Select DriveLetter, FileSystemLabel, HealthStatus, Size, SizeRemaining


# Prompt to checkdisk on local volumes, skipping volumes with no drive letter (System Reserved)
$confirmation = Read-Host "Do you want to run Chkdsk on attached volumes? (y/n)"
if ($confirmation -eq 'y') {
  Write-Host "`n"
    Write-Host "Running Chkdsk against local volumes..."
    Write-Host "`n"
    Get-Volume | Where-Object { ($_.DriveType -eq "Fixed") -and ($_.DriveLetter -ne $NULL)} | foreach { repair-volume -driveletter $_.DriveLetter -Scan}
}
else {
    Write-Host "Skipping Chkdsk"
}


# Check resource usage
Write-Host "`n"
Write-Host "Checking resource utilization..."
Write-Host "`n"
$Counters = @(
    '\Processor(_Total)\% Processor Time',
    '\Memory\Available MBytes',
    '\Memory\cache faults/sec',
    '\physicaldisk(_total)\% disk time',
    '\physicaldisk(_total)\current disk queue length',
    '\Paging File(_total)\% Usage'
)
Get-Counter -Counter $Counters -MaxSamples 1 | ForEach {
    $_.CounterSamples | ForEach {
        [pscustomobject]@{
            TimeStamp = $_.TimeStamp
            Path = $_.Path
            Value = $_.CookedValue
        }
    }
} 



#Manual
#Check antivirus log
#RAID and virtual disk check
#Firmware update
