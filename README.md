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

}

Invoke-ADGroupMembershipMonitoring @paramADGroupMembershipMonitoring
```

Scheduled task:
```
"C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe" -File "C:\Scripts\Run-ADGroupMembershipMonitoring.ps1"
```

## Features

### The first time you run the script
The function is creating folders if they not exists:
 * **Current** directory: the current AD group membership is queried and saved in the file (it won't touch the file if it's the same membership) on each run of the function.
 * **History** directory: contains the list of changes from the past. One file per AD Group per domain, if multiple changes occur, the function will append the change in the same file.


### Comparing
The membership of each group is saved in a CSV file "DOMAIN_GROUPNAME-membership.csv"
If the file does not exist, the script will create one, so the next time it will be able to compare the  membership with this file.

### Change History
Each time a change is detected (Add or Remove an Account (Nested or Not)) a CSV file will be generated with the following name: "DOMAIN_GROUPNAME-ChangesHistory-yyyyMMdd-hhmmss.csv"

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

``` powershell
PS> Invoke-ADGroupMembershipMonitoring -Group "FXGroup01","FXGroup02" -Emailfrom Reporting@fx.lab -Emailto "Catfx@fx.lab" -EmailServer 192.168.1.10 -Verbose
```

## Change Log
[Change notes for each release](CHANGELOG.md)