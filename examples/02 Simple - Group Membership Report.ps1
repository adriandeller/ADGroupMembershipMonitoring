break

Import-Module -Name ADGroupMembershipMonitoring

$paramADGroupMembershipMonitoring = @{
    Group               = 'Domain Admins'
    Recursive           = $true
    EmailSubjectPrefix  = '[Group Membership Monitoring]'
    EmailTo             = 'it-department@company.com'
    EmailFrom           = 'noreply@company.com'
    EmailServer         = 'mail.company.com'
    EmailEncoding       = 'UTF8'
    SendEmail           = $true
    ForceAction         = $true
    IncludeMembers      = $true
    ExcludeSummary      = $true
    ExcludeChanges      = $true
    ExcludeHistory      = $true
    Path                = 'C:\Scripts\ADGroupMembershipMonitoring'
}

Invoke-ADGroupMembershipMonitoring @paramADGroupMembershipMonitoring
