# Check if the script is running with administrative privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch the script with administrative privileges
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-File `"$PSCommandPath`""
    exit
}

$WarningPreference = 'SilentlyContinue'
$ErrorActionPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'
$InformationPreference = 'Continue'


# Stop the print spooler service and wait
Write-Host "Stopping the print spooler service..."
Stop-Service -Name Spooler -Force
Start-Sleep -Seconds 2

# Kill the spooler.exe process and printpipelinesvc.exe process
Write-Host "Killing spool adjacent processes..."
Get-Process -Name spoolsv, PrintIsolationHost, Printfilterpipelinesvc | Stop-Process -Force
Start-Sleep -Seconds 5

# Clear all print jobs by deleting files from the PRINTERS folder
$printersFolder = "$env:SystemRoot\System32\spool\PRINTERS"
if (Test-Path -Path $printersFolder) {
    Write-Host "Deleting print job files from the PRINTERS folder..."
    Get-ChildItem -Path $printersFolder | Remove-Item -Force
}
else {
    Write-Host "PRINTERS folder not found. No print jobs to clear."
}

# Check if the print spooler service is running
$spoolerService.Refresh()
if ($spoolerService.Status -ne 'Running') {
    # Start the print spooler service
    Write-Host "Starting the print spooler service, please wait..."
    Start-Service -Name Spooler
}