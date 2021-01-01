@{
    Email = @{
            EmailSubjectPrefix = '[AD Group Membership Monitoring]'
    }
    Folder = @{
        Current   = 'Current'
        History   = 'History'
        HTML      = 'HTML'
        OneReport = 'OneReport'
    }
    CSV = @{
        ChangeHistoryProperty = 'DateTime', 'State', 'Name', 'SamAccountName', 'ObjectClass'
        MembershipProperty    = 'Name', 'SamAccountName', 'ObjectClass'
    }
}
