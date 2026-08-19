# Run this in an ELEVATED PowerShell (Right-click PowerShell -> Run as Administrator)
# Uses schtasks.exe (not the ScheduledTasks PowerShell module) to avoid a known
# PowerShell bug where Set-ScheduledTask spuriously fails with "user name or
# password is incorrect" on tasks that use saved-password logon, even when no
# credentials are being changed.
#
# The /TR value must literally contain quote characters around the .bat path
# (Task Scheduler's own requirement for paths with spaces) - building it as a
# single-quoted PowerShell string avoids PowerShell's own escaping rules
# mangling those embedded quotes.
$tr = '"C:\Users\user\Desktop\python\Libyana Daily Report - Copy\run_scheduler.bat"'
schtasks /Change /TN "Libyana NPM Daily Scheduler" /TR $tr

Write-Output "--- Updated task ---"
schtasks /Query /TN "Libyana NPM Daily Scheduler" /V /FO LIST | Select-String "Task To Run|Start In"
