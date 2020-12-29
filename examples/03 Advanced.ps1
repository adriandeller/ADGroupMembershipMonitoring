break

Import-Module -Name ADGroupMembershipMonitoring

$paramADGroupMembershipMonitoring = @{
    SearchRoot          = 'OU=Groups,DC=company,DC=com'
    GroupScope          = 'Universal'
    GroupFilter         = "name -like 'it-role-*'"
    Recursive           = $true
    EmailSubjectPrefix  = '[High Privileged Groups]'
    EmailTo             = 'it-department@company.com'
    EmailFrom           = 'noreply@company.com'
    EmailServer         = 'mail.company.com'
    EmailEncoding       = 'UTF8'
    SaveAsHTML          = $true
    Path                = 'C:\Scripts\ADGroupMembershipMonitoring\HighPrivilegedGroups'
}

Invoke-ADGroupMembershipMonitoring @paramADGroupMembershipMonitoring
