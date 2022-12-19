$ScriptPath = $env:ProgramFiles +"\YourPath\Scripts\Set-TeamsFirewallRules"

Write-Output "Removing $($ScriptPath)"

Remove-Item $ScriptPath -Recurse

Write-Output "Unregistering Task Set-TeamsFirewallRules"

Unregister-ScheduledTask -TaskName "Set-TeamsFirewallRules" -Confirm:$false

Write-Output "Removing firewall rules related to Teams Access"
$TeamsBlock = Get-NetFirewallRule -Name *Teams* | Where-Object {$_.Enabled -eq 'True'}
If ($null -ne $TeamsBlock)
{
    Foreach ($Rule in $TeamsBlock)
    {
        Write-Output "Disabling blocking firewall rule $($Rule.DisplayName)"
        Remove-NetFirewallRule $Rule.Name -Confirm:$false
    }
}

Stop-Transcript
