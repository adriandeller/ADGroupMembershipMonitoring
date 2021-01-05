# ADGroupMembershipMonitoring

This PowerShell module provides functions for monitoring Active Directory groups, tracking changes and send e-mail notifications on changes.

## Credits

The modules functionality is based mainly on the code of the PowerShell script [Monitor-ADGroupMembership](https://github.com/lazywinadmin/Monitor-ADGroupMembership) by [@lazywinadmin](https://twitter.com/lazywinadmin).

## Installation

#### Download from PowerShell Gallery (PowerShell v5+)

You can install the script directly from the PowerShell Gallery.

``` powershell
Install-Module -Name ADGroupMembershipMonitoring
```

## Schedule the script

On frequent question I get for this script is how to use the Task Scheduler to run this script.

The recommended way to do this, is creating a wrapper script which is then called by the scheduled task.

``` powershell
Import-Module -Name ADGroupMembershipMonitoring

$paramADGroupMembershipMonitoring = @{
    SearchRoot          = 'OU=Groups,DC=company,DC=com'
    GroupScope          = 'Universal'
    GroupFilter         = "name -like 'IT-Role-*'"
    Recursive           = $true
    EmailSubjectPrefix  = '[High Privileged Groups]'
    EmailTo             = 'it-department@company.com'
    EmailFrom           = 'noreply@company.com'
    EmailServer         = 'mail.company.com'
    SendEmail           = $true
    SaveReport          = $true
    Path                = 'C:\Scripts\ADGroupMembershipMonitoring\HighPrivilegedGroups'
}

Invoke-ADGroupMembershipMonitoring @paramADGroupMembershipMonitoring
```

Scheduled task:
```
"C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -NoProfile -NonInteractive -WindowStyle Hidden -File 'C:\Scripts\ADGroupMembershipMonitoring\Invoke-ADGroupMembershipMonitoring-HPG.ps1'
```

## Features

### Configuration (Defaults)
Default values for common settings are saved in a PSD1 file in the module directory.
``` powershell
.\config\configuration.psd1
```

### Data storage
The function will create folders if they not exists:
 * **Current** directory: the current AD group membership is queried and saved in the file (it won't touch the file if it's the same membership) on each run of the function.
 * **History** directory: contains the list of changes from the past. One file per AD Group per domain, if multiple changes occur, the function will append the change in the same file.
 * **HTML** directory:
 * **OneReport** directory:


### Comparing
The membership of each group is saved in a CSV file "DOMAIN_GROUPNAME-membership.csv"
If the file does not exist, the script will create one, so the next time it will be able to compare the  membership with this file.

### Change History
Each time a change is detected (Add or Remove an Account (Nested or Not)) a CSV file will be generated with the following name: "DOMAIN_GROUPNAME-ChangeHistory.csv"

When generating the HTML Report, the script will add this Change History to the Email (if there is one to add)

### Reporting
A report generated when a change is detected.
Also, if some Change History files for this group exists, it will be added to the report.
Finally at the end of the report, information on when, where and who ran the script.


## Requirements
* Read Permission in Active Directory on the monitored groups
* PowerShell Module ActiveDirectory (RSAT)
* optional: Scheduled Task in order to check every X seconds/minutes/hours

## Examples

This will query the groups 'Domain Admins' and 'Enterprise Admins' and send an email to 'To@Company.com' using the address 'From@Company.com' and the server 'mail.company.com'.
All data, CSV and HTML report files, are saved in subfolders in the directory 'C:\ADGroupMembershipMonitoringData'.
Additionally the 'Verbose' switch is enabled to show the activities of the PowerShell function.
``` powershell
PS> Invoke-ADGroupMembershipMonitoring -Group 'Domain Admins','Enterprise Admins' -EmailFrom 'From@Company.com' -EmailTo 'To@Company.com' -EmailServer 'mail.company.com' -Path 'C:\ADGroupMembershipMonitoringData' -Verbose
```

This will query the group 'Domain Admins' recursively and send an email to 'To@Company.com' using the address 'From@Company.com' and the server 'mail.company.com'.
with the 'Recursive' switch, group members with indirect membership (through group nesting) will also be searched for.
``` powershell
PS> Invoke-ADGroupMembershipMonitoring -Group 'Domain Admins' -Recursive -EmailFrom 'From@Company.com' -EmailTo 'To@Company.com' -EmailServer 'mail.company.com' -Path 'C:\ADGroupMembershipMonitoringData'
```

This will query all the groups present in the CanonicalName 'Company.com/Test/Groups' and send an email using the encoding 'UTF8' to 'To@Company.com' using the address'From@Company.com' and the server 'mail.company.com'.
``` powershell
PS> Invoke-ADGroupMembershipMonitoring -SearchRoot 'Company.com/Test/Groups' -EmailEncoding 'UTF8' -EmailFrom 'From@Company.com' -EmailTo 'To@Company.com' -EmailServer 'mail.company.com' -Path 'C:\ADGroupMembershipMonitoringData'
```

This will query all the groups present in the file 'ListOfHighPrivilegedGroups.txt' and send an email to 'To@Company.com' using the address'From@Company.com' and the server 'mail.company.com'.
``` powershell
PS> Invoke-ADGroupMembershipMonitoring -File .\ListOfHighPrivilegedGroups.txt -EmailFrom 'From@Company.com' -EmailTo 'To@Company.com' -EmailServer 'mail.company.com' -Path 'C:\ADGroupMembershipMonitoringData'
```

This will query all the groups present in the file 'ListOfHighPrivilegedGroups.txt' against the Domain Controller 'dc01.company.com' and send an email to 'To@Company.com' using the address'From@Company.com' and the server 'mail.company.com'.
``` powershell
PS> Invoke-ADGroupMembershipMonitoring -File .\ListOfHighPrivilegedGroups.txt -Server 'dc01.company.com' -EmailFrom 'From@Company.com' -EmailTo 'To@Company.com' -EmailServer 'mail.company.com' -Path 'C:\ADGroupMembershipMonitoringData'
```

## Change Log
[Change notes for each release](CHANGELOG.md)
