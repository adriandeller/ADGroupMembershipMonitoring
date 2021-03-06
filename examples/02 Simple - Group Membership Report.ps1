break

Import-Module -Name ADGroupMembershipMonitoring

$paramADGroupMembershipMonitoring = @{
    Path                = 'C:\Scripts\ADGroupMembershipMonitoring'
    Group               = 'Domain Admins'
    Recursive           = $true
    EmailSubjectPrefix  = '[Group Membership Reporting]'
    EmailTo             = 'it-department@company.com'
    EmailFrom           = 'noreply@company.com'
    EmailServer         = 'mail.company.com'
    SendEmail           = $true
    ForceAction         = $true
    IncludeMembers      = $true
    ExcludeSummary      = $true
    ExcludeChanges      = $true
    ExcludeHistory      = $true
}

Invoke-ADGroupMembershipMonitoring @paramADGroupMembershipMonitoring
