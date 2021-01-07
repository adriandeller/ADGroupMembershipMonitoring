break

Import-Module -Name ADGroupMembershipMonitoring

$paramADGroupMembershipMonitoring = @{
    Path                = 'C:\Scripts\ADGroupMembershipMonitoring'
    Group               = 'Domain Admins'
    Recursive           = $true
    EmailSubjectPrefix  = '[AD Group Membership Monitoring]'
    EmailTo             = 'it-department@company.com'
    EmailFrom           = 'noreply@company.com'
    EmailServer         = 'mail.company.com'
    SaveReport          = $true
    IncludeMembers      = $true
}

Invoke-ADGroupMembershipMonitoring @paramADGroupMembershipMonitoring
