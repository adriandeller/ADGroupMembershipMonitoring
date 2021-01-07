break

Import-Module -Name ADGroupMembershipMonitoring

$paramADGroupMembershipMonitoring = @{
    Path                = 'C:\Scripts\ADGroupMembershipMonitoring\HighPrivilegedGroups'
    SearchRoot          = 'OU=Groups,DC=company,DC=com'
    GroupScope          = 'Universal'
    GroupFilter         = "name -like 'it-role-*'"
    Recursive           = $true
    EmailSubjectPrefix  = '[High Privileged Groups]'
    EmailTo             = 'it-department@company.com'
    EmailFrom           = 'noreply@company.com'
    EmailServer         = 'mail.company.com'
    SaveReport          = $true
    ForceAction         = $true
}

Invoke-ADGroupMembershipMonitoring @paramADGroupMembershipMonitoring
