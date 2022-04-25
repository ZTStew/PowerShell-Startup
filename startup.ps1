<#
    File: startup.ps1
    Task:
        Starts up applications used every day in the morning
        Closes unnessisary applications that open on boot
#>

# Start-Process "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
# Start-Process "C:\Users\zachary.stewart\AppData\Local\slack\slack.exe"

function Find-Open {
  param (
    $applicationName,
    $applicationPath
  )
  $application = Get-Process $applicationName -ErrorAction SilentlyContinue
  # gets path of process
  # $processPath = Get-Process $applicationName | Select-Object -ExpandProperty Path
  # Write-Output $processPath 
  # Start-Process $processPath 

  if (!($application)) {
    Write-Output $applicationName" not found"
    Start-Process $applicationPath
    # upper cases application name
    Write-Output $("Starting: " + ( Get-Culture ).TextInfo.ToTitleCase( $applicationName.ToLower() ))
    # return $false
  } else {
    Write-Output $applicationName" found"
    # return $true
  }
}
# Get-Process | Select-Object -ExpandProperty Path

Find-Open -applicationName "outlook" -applicationPath "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
Find-Open -applicationName "slack" -applicationPath "C:\Users\zachary.stewart\AppData\Local\slack\slack.exe"


# Checks if Outlook process exists, if not, open Outlook
# $outlook = Get-Process outlook -ErrorAction SilentlyContinue
# if (!($outlook)) {
#     Write-Output "Outlook closed"
# } else {
#     Write-Output "Outlook Open"
# }
