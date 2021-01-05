break

Import-Module -Name ADGroupMembershipMonitoring

$paramADGroupMembershipMonitoring = @{
    Group               = 'Domain Admins'
    Recursive           = $true
    EmailSubjectPrefix  = '[AD Group Membership Monitoring]'
    EmailTo             = 'it-department@company.com'
    EmailFrom           = 'noreply@company.com'
    EmailServer         = 'mail.company.com'
    EmailEncoding       = 'UTF8'
    SaveReport          = $true
    IncludeMembers      = $true
    Path                = 'C:\Scripts\ADGroupMembershipMonitoring'
}

Invoke-ADGroupMembershipMonitoring @paramADGroupMembershipMonitoring
