break

Import-Module -Name ADGroupMembershipMonitoring

$paramADGroupMembershipMonitoring = @{
    Path                = 'C:\Scripts\ADGroupMembershipMonitoring\HighPrivilegedGroups'
    LDAPFilter          = '(|(adminCount=1)(Name=DnsAdmins))'
    Recursive           = $true
    EmailSubjectPrefix  = '[High Privileged Groups]'
    EmailTo             = 'it-department@company.com'
    EmailFrom           = 'noreply@company.com'
    EmailServer         = 'mail.company.com'
    IncludeMembers      = $true
    SaveReport          = $true
    SendEmail           = $true
    ForceAction         = $true
}

Invoke-ADGroupMembershipMonitoring @paramADGroupMembershipMonitoring
