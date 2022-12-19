$LogPath = 'C:\Logs'
$LogName = $LogPath + "\Setup-TeamsFirewallRule.log"

If ((Test-Path $LogPath) -eq $false)
{
    New-Item -Path c:\Logs -ItemType Directory
}

Start-Transcript -Path $LogName -Append

$ScriptPath = $env:ProgramFiles +"\YourPath\Scripts\Set-TeamsFirewallRules"

Write-Output "Checking for script paths"
If ((Test-Path $ScriptPath) -eq $false)
{
    Write-Output "Script path did not exist, building it now $($ScriptPath)"
    New-Item $ScriptPath -ItemType Directory 
}

Write-Output "Copying Set-TeamsFirewallRules and Uninstall scripts to script path"

$Files = Get-ChildItem $PSScriptroot

foreach ($File in $Files)
{
    Write-Output "Copying $($file.Name) to $($ScriptPath)"
    Copy-Item -Path $File.FullName -Destination $ScriptPath
}

$ScriptLaunch = $ScriptPath + "\Set-TeamsFirewallRules.ps1"

#Check for existing task and remove it if it's there
If ((Get-ScheduledTask 'Set-TeamsFirewallRules' -ErrorAction SilentlyContinue))
{
    Write-Output "Found existing task, removing it now..."
    Unregister-ScheduledTask 'Set-TeamsFirewallRules' -Confirm:$false
}

Write-Output "Building scheduled task to run the script at logon for each user"
$Action = New-ScheduledTaskAction -Execute powershell.exe -Argument '-ExecutionPolicy Bypass -File "C:\Program Files (x86)\YourPath\Scripts\Set-TeamsFirewallRules\Set-TeamsFirewallRules.ps1"'
$Trigger = New-ScheduledTaskTrigger -AtLogOn -RandomDelay (New-TimeSpan -Minutes 1)
$Principal = New-ScheduledTaskPrincipal "SYSTEM" -RunLevel Highest
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries 
$Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings -Description "This will create new Teams firewall rules for each user at logon" 
Register-ScheduledTask 'Set-TeamsFirewallRules' -InputObject $Task

Write-Output "Initiating first run of the Set-TeamsFirewallRules script"
Start-ScheduledTask -TaskName "Set-TeamsFirewallRules" 

Stop-Transcript
