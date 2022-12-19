$LogPath = 'C:\Logs'
$LogName = $LogPath + "\Set-TeamsFirewallRule.log"

If ((Test-Path $LogPath) -eq $false)
{
    New-Item -Path c:\Logs -ItemType Directory
}

Start-Transcript -Path $LogName -Force

#Get all the user profiles that are on the system
$users = Get-ChildItem (Join-Path -Path $env:SystemDrive -ChildPath 'Users') -Exclude 'Public', 'ADMIN*','Help*','svc*' 

#Generate a date code for accounts that have had activity in the last 180 days
$StaleAccountDate = (Get-Date).AddDays(-180)


if ($null -ne $users) 
    {

    Write-Output "Checking for firewall rules that block Teams Access"
    $TeamsBlock = Get-NetFirewallRule -Name *Teams* | Where-Object {($_.action -eq "Block") -and ($_.Enabled -eq 'True')}
    If ($null -ne $TeamsBlock)
    {
        Foreach ($Rule in $TeamsBlock)
        {
            Write-Output "Disabling blocking firewall rule $($Rule.DisplayName)"
            Set-NetFirewallRule $Rule.Name -Enabled False
        }
    }

    foreach ($user in $users) 
        {
            #Check the last time something was written to the file path for the user we are checking
            $LastAccessTime = Get-ChildItem $user.FullName | sort lastwritetime -Descending

            If ($LastAccessTime[0].LastWriteTime -le $StaleAccountDate)
            {
                Write-Output "Skipping $($user) because the last access time on the account is too old"
            }

            If ($LastAccessTime[0].LastWriteTime -ge $StaleAccountDate)
            {
                Write-Output "Checking $($user.FullName)"
                #Generate the path that will be checked
                $progPath = Join-Path -Path $user.FullName -ChildPath "AppData\Local\Microsoft\Teams\Current\Teams.exe"
                    
                    if (Test-Path $progPath) 
                    {
                        #Check that the firewall rule already exists, if it does skip, if not, add it
                        $RuleCheck = Get-NetFirewallApplicationFilter -Program $progPath -ErrorAction SilentlyContinue

                        if ($Null -eq $RuleCheck) 
                        {
                            Write-Output "Creating new firewall rules for $($user)"
                            $ruleName = "Teams.exe for user $($user.Name)"
                            "UDP", "TCP" | ForEach-Object { New-NetFirewallRule -Name ($RuleName + " " + $_.ToString()) -DisplayName ($RuleName + " " + $_.ToString()) -Direction Inbound -Profile Any -Program $progPath -Action Allow -Protocol $_ }
                            Clear-Variable ruleName
                        }

                        If ($Null -ne $RuleCheck)
                        {
                            Write-Output "Looks like there are already rules in place for $($user) so we will not add any"
                        }
                    }

                    if ((Test-Path $progPath) -eq $false)
                    {
                        Write-Output "Looks like $($user) has not launched Teams, we will be skipping"
                    }
            }
        #Clear the variables for the next user run
        Clear-Variable TeamsBlock
        Clear-Variable RuleCheck
        Clear-Variable progPath
            
    }
}


Stop-Transcript
