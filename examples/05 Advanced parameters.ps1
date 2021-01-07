break

Import-Module -Name ADGroupMembershipMonitoring

$paramADGroupMembershipMonitoring = @{
    Path                = 'C:\Scripts\ADGroupMembershipMonitoring\HighPrivilegedGroups'
    Group               = 'hr-department'
    Recursive           = $true
    EmailSubjectPrefix  = '[HR Group Membership Monitoring]'
    #EmailTo             = 'it-department@company.com'
    EmailToManger       = $true
    EmailToSelf         = $true
    EmailFrom           = 'noreply@company.com'
    EmailServer         = 'mail.company.com'
    SendEmail           = $true
    IncludeMembers      = $true
    ExcludeSummary      = $true
    ExcludeChanges      = $true
    ExcludeHistory      = $true
}

Invoke-ADGroupMembershipMonitoring @paramADGroupMembershipMonitoring
