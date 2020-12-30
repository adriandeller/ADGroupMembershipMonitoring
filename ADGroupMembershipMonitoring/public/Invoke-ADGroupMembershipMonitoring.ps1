function Invoke-ADGroupMembershipMonitoring
{
    <#
    .DESCRIPTION
        This function is monitoring group(s) in Active Directory and sends an email when someone is changing the membership.
        It will also report the Change History made for this/those group(s).
    .SYNOPSIS
        This function is monitoring group(s) in Active Directory and sends an email when someone is changing the membership.
    .PARAMETER Group
        Specify the group(s) to query in Active Directory.
        You can also specify the 'DN','GUID','SID' or the 'Name' of your group(s).
        Using 'Domain\Name' will also work.
    .PARAMETER Recursive
        by default, search for group members with direct membership,
        Specify this switch and group members with indirect membership through group nesting will also be searched
    .PARAMETER SearchRoot
        Specify the DN, GUID or canonical name of the domain or container to search. By default, the script searches the entire sub-tree of which SearchRoot is the topmost object (sub-tree search). This default behavior can be altered by using the SearchScope parameter.
    .PARAMETER SearchScope
        Specify one of these parameter values
            'Base' Limits the search to the base (SearchRoot) object.
                The result contains a maximum of one object.
            'OneLevel' Searches the immediate child objects of the base (SearchRoot)
                object, excluding the base object.
            'Subtree'   Searches the whole sub-tree, including the base (SearchRoot)
                object and all its child objects.
    .PARAMETER GroupScope
        Specify the group scope of groups you want to find. Acceptable values are:
            'Global';
            'Universal';
            'DomainLocal'.
    .PARAMETER GroupType
        Specify the group type of groups you want to find. Acceptable values are:
            'Security'
            'Distribution'.
    .PARAMETER GroupFilter
        Specify the filter you want to use to find groups
    .PARAMETER File
        Specify the File where the Group are listed. DN, SID, GUID, or Domain\Name of the group are accepted.
    .PARAMETER EmailServer
        Specify the Email Server IPAddress/FQDN.
    .PARAMETER EmailPort
        Specify the port for the Email Server
    .PARAMETER EmailTo
        Specify the Email Address(es) of the Destination. Example: fxcat@fx.lab
    .PARAMETER EmailFrom
        Specify the Email Address of the Sender. Example: Reporting@fx.lab
    .PARAMETER EmailEncoding
        Specify the Body and Subject Encoding to use in the Email.
        Default is ASCII.
    .PARAMETER EmailSubjectPrefix
        Specify a prefix for the E-Mail subject
    .PARAMETER Server
        Specify the Domain Controller to use.
        Aliases: DomainController, Service
    .PARAMETER SaveAsHTML
        Specify if you want to save a local copy of the Report.
        It will be saved under the directory "HTML".
    .PARAMETER IncludeMembers
        Specify if you want to include all members in the Report.
    .PARAMETER AlwaysReport
        Specify if you want to generate a Report each time.
    .PARAMETER OneReport
        Specify if you want to only send one email with all group report as attachment
    .PARAMETER ExtendedProperty
        Specify if you want to add Enabled and PasswordExpired Attribute on members in the report
    .PARAMETER EmailCredential
        Specify alternative credential to use. By default it will use the current context account.
    .PARAMETER Path
        Specify a path for data storage, where subfolders will be created for the CSV and HTML files
    .EXAMPLE
        PS> Invoke-ADGroupMembershipMonitoring -Group 'Domain Admins' -EmailFrom 'From@Company.com' -EmailTo 'To@Company.com' -EmailServer 'mail.company.com'
        This will query the group  'Domain Admins' and send an email to 'To@Company.com' using the address 'From@Company.com' and the server 'mail.company.com'.
    .NOTES
        Author:      Adrian Deller
        Contact:     adrian.deller@unibas.ch
        Created:     2019-10-05
        Updated:     2020-12-29
        Version:     0.2.0
    .ExternalHelp ADGroupMembershipMonitoring-help.xml
    #>

    [CmdletBinding(DefaultParameterSetName = 'Group')]

    param
    (
        [Parameter(ParameterSetName = 'Group', Mandatory, HelpMessage = 'You must specify at least one Active Directory group')]
        [ValidateNotNull()]
        [Alias('DN', 'DistinguishedName', 'GUID', 'SID', 'Name')]
        [string[]]
        $Group,

        [Parameter(HelpMessage = 'Should the AD group members be searched recursively?')]
        [switch]
        $Recursive,

        [Parameter(ParameterSetName = 'ADFilter', Mandatory, HelpMessage = 'You must specify at least one Active Directory OU')]
        [Alias('SearchBase')]
        [string[]]
        $SearchRoot,

        [Parameter(ParameterSetName = 'ADFilter')]
        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [string]
        $SearchScope,

        [Parameter(ParameterSetName = 'ADFilter')]
        [ValidateSet('Global', 'Universal', 'DomainLocal')]
        [String]
        $GroupScope,

        [Parameter(ParameterSetName = 'ADFilter')]
        [ValidateSet('Security', 'Distribution')]
        [string]
        $GroupType,

        [Parameter(ParameterSetName = 'ADFilter')]
        [string]
        $GroupFilter,

        [Parameter(ParameterSetName = 'File', Mandatory, HelpMessage = 'You must specify at least one file')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {
                if (-not (Test-Path -Path $_) )
                {
                    throw "File '$_' does not exist"
                }
                if (-not (Test-Path -Path $_ -PathType Leaf) )
                {
                    throw "The File argument must be a file"
                }
                return $true
            })]
        [string[]]
        $File,

        [Parameter()]
        [Alias('DomainController', 'Service')]
        [string]
        $Server,

        [Parameter(Mandatory, HelpMessage = 'You must specify the sender E-Mail Address')]
        [ValidatePattern("[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?")]
        [string]
        $EmailFrom,

        [Parameter(Mandatory, HelpMessage = 'You must specify the destination E-Mail Address')]
        [ValidatePattern("[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?")]
        [string[]]
        $EmailTo,

        [Parameter(Mandatory, HelpMessage = 'You must specify the Mail Server to use (FQDN or ip address)')]
        [string]
        $EmailServer,

        [Parameter(HelpMessage = 'You can specify an alternate port on the (SMTP) Mail Server')]
        [string]
        $EmailPort = '25',

        [Parameter(HelpMessage = 'You can provide a prefix for the E-Mail subject')]
        [string]
        $EmailSubjectPrefix,

        [Parameter(HelpMessage = 'You can specify the type of encoding')]
        [ValidateSet('ASCII', 'UTF8', 'UTF7', 'UTF32', 'Unicode', 'BigEndianUnicode', 'Default')]
        [string]
        $EmailEncoding = 'Default',

        [Parameter(HelpMessage = 'You can provide credentials for the Mail Server')]
        [System.Management.Automation.PSCredential]
        $EmailCredential = [System.Management.Automation.PSCredential]::Empty,

        [Parameter()]
        [Switch]
        $SaveAsHTML,

        [Parameter()]
        [Switch]
        $IncludeMembers,

        [Parameter()]
        [Switch]
        $AlwaysReport,

        [Parameter()]
        [Switch]
        $OneReport,

        [Parameter()]
        [Switch]
        $ExtendedProperty,

        [Parameter(Mandatory, HelpMessage = 'You must specify a path for data storage')]
        [Alias('OutputPath', 'FolderPath')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Path
    )

    Begin
    {
        $ScriptName = $MyInvocation.MyCommand.Name

        Write-Verbose -Message "[$ScriptName][Begin]"

        $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()

        $GroupList = New-Object System.Collections.Generic.List[Object]

        try
        {
            #Region Import Modules
            $RequiredModules = 'ActiveDirectory'

            foreach ($ModuleName in $RequiredModules)
            {
                if (Get-Module -ListAvailable -Name $ModuleName -Verbose:$false)
                {
                    Import-Module -Name $ModuleName -Verbose:$false
                    Write-Verbose -Message "    [+] Imported module '$ModuleName'"
                }
                else
                {
                    Write-Warning "    [-] Module '$ModuleName' does not exist"
                    return
                }
            }
            #EndRegion

            #Region Configuration
            # Import default values from config file
            $Configuration = Get-Configuration -Verbose
            Write-Verbose -Message "    [+] Imported configuration"

            $PSDefaultParameterValues = @{
                'Convert-DateTimeString:Verbose' = $false
                'Get-ADComputer:Property' = 'DisplayName', 'PasswordExpired'
                'Get-ADUser:Property' = 'DisplayName', 'PasswordExpired'
            }

            # CSV columns
            $ChangeHistoryCsvProperty = 'DateTime', 'State', 'DisplayName', 'Name', 'SamAccountName', 'DistinguishedName'

            # Report table columns
            $MembershipChangeTableProperty = 'DateTime', 'State', 'Name', 'SamAccountName', 'DistinguishedName'
            $ChangeHistoryTableProperty = 'DateTime', 'State', 'Name', 'SamAccountName', 'DistinguishedName'
            $MembersTableProperty = 'Name', 'SamAccountName', 'ObjectClass'

            if ($PSBoundParameters['ExtendedProperty'])
            {
                $MembersTableProperty += 'Enabled', 'PasswordExpired', 'DistinguishedName'
            }

            if ($PSBoundParameters['EmailSubjectPrefix'])
            {
                $EmailSubjectPrefix = $PSBoundParameters['EmailSubjectPrefix']
            }
            else
            {
                $EmailSubjectPrefix = $Configuration.Email.EmailSubjectPrefix
            }

            # Set the Date and Time formats
            $ChangesDateTimeFormat = 's'                         # format for export to CSV files
            $ReporCreatedDateTimeFormat = 'dd.MM.yyyy HH:mm:ss'  # format for report creation date/time information
            $ReportChangesDateTimeFormat = 'yyyy-MM-dd HH:mm:ss' # format for DateTime property in HTML reports
            $FileNameDateTimeFormat = 'yyyyMMdd_HHmmss'          # format for DateTime information in CSV file names

            # list of known DateTime formats used in legacy CSV files
            $KnownInputFormat = 'yyyyMMdd-HH:mm:ss', $ChangesDateTimeFormat

            # HTML Report Settings
            $Report = "<p style=`"background-color:white; font-family:Calibri; font-size:10pt`">" +
            "<strong>Report created:</strong> $(Get-Date -Format $ReporCreatedDateTimeFormat)<br>" +
            "<strong>Account:</strong> $env:USERDOMAIN\$($env:USERNAME) on $($env:COMPUTERNAME)" +
            "</p>"

            $Head = "<style>" +
            "body {background-color:white; font-family:Calibri; font-size:11pt}" +
            "h2,h3 {font-family:Calibri}" +
            "table {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse}" +
            "th {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:`"#00297A`";font-color:white}" +
            "td {border-width: 1px;padding-right: 2px;padding-left: 2px;padding-top: 0px;padding-bottom: 0px;border-style: solid;border-color: black;background-color:white}" +
            "</style>"

            $Head2 = "<style>" +
            "body {background-color:white; font-family:Calibri; font-size:10pt;}" +
            "h2,h3 {font-family:Calibri}" +
            "table {border-width:1px; border-style:solid; border-color:black; border-collapse:collapse;}" +
            "th {border-width:1px; padding:3px; border-style:solid; border-color:black; background-color:`"#C0C0C0`"}" +
            "td {border-width:1px; padding-right:2px; padding-left:2px; padding-top:0px; padding-bottom:0px; border-style:solid; border-color:black; background-color:white}" +
            "</style>"
            #EndRegion

            #Region Folders for data storage
            # Create the folders if not present
            if (-not (Test-Path -Path $Path))
            {
                $null = New-Item -Path $Path -ItemType Directory -ErrorAction Stop

                if (-not (Test-Path -Path $Path -PathType Container) )
                {
                    throw "The folder '$Path' does not exist"
                }

                Write-Verbose -Message "    [+] Created folder for data storage '$Path'"
            }
            else
            {
                Write-Verbose -Message "    [i] Folder for data storage exists: '$Path'"
            }

            $Subfolders = $Configuration.Folder.Keys

            foreach ($Subfolder in $Subfolders)
            {
                $SubfolderPath = $Path + "\{0}" -f $Configuration.Folder[$Subfolder]

                New-Variable -Name ('ScriptPath{0}' -f $Subfolder) -Value $SubfolderPath

                if (-not (Test-Path -Path $SubfolderPath))
                {
                    $null = New-Item -Path $SubfolderPath -ItemType Directory -ErrorAction Stop

                    if (-not (Test-Path -Path $SubfolderPath -PathType Container) )
                    {
                        throw "The folder '$SubfolderPath' does not exist"
                    }

                    Write-Verbose -Message "    [+] Created subfolder '$SubfolderPath'"
                }
                else
                {
                    Write-Verbose -Message "    [i] Subfolder exists: '$SubfolderPath'"
                }
            }
            #EndRegion
        }
        catch
        {
            Write-Warning -Message "[$ScriptName][Begin] Something went wrong"
            throw $_.Exception.Message
        }
    }
    Process
    {
        try
        {
            switch ($PSCmdlet.ParameterSetName)
            {
                'ADFilter'
                {
                    Write-Verbose -Message "[*] Get AD Groups using AD filter"

                    foreach ($SearchRootItem in $SearchRoot)
                    {
                        # ADGroup Splatting
                        $ADGroupParams = @{ }

                        $ADGroupParams.SearchBase = $SearchRootItem

                        # Server parameter specified
                        if ($PSBoundParameters['Server'])
                        {
                            Write-Verbose -Message "[i] Server specified"
                            $ADGroupParams.Server = $Server
                        }

                        # SearchScope parameter specified
                        if ($PSBoundParameters['SearchScope'])
                        {
                            Write-Verbose -Message "[i] SearchScope specified"
                            $ADGroupParams.SearchScope = $SearchScope
                        }

                        # GroupScope parameter specified
                        if ($PSBoundParameters['GroupScope'])
                        {
                            Write-Verbose -Message "[i] GroupScope specified"
                            $ADGroupParams.Filter = "GroupScope -eq `'$GroupScope`'"
                        }

                        # GroupType parameter specified
                        if ($PSBoundParameters['GroupType'])
                        {
                            Write-Verbose -Message "[i] GroupType specified"

                            if ($ADGroupParams.Filter)
                            {
                                $ADGroupParams.Filter = "$($ADGroupParams.Filter) -and GroupCategory -eq `'$GroupType`'"
                            }
                            else
                            {
                                $ADGroupParams.Filter = "GroupCategory -eq '$GroupType'"
                            }
                        }

                        # GroupFilter parameter specified
                        if ($PSBoundParameters['GroupFilter'])
                        {
                            Write-Verbose -Message "[i] GroupFilter specified"

                            if ($ADGroupParams.Filter)
                            {
                                $ADGroupParams.Filter = "$($ADGroupParams.Filter) -and `'$GroupFilter`'"
                            }
                            else
                            {
                                $ADGroupParams.Filter = $GroupFilter
                            }
                        }

                        if (-not ($ADGroupParams.Filter))
                        {
                            $ADGroupParams.Filter = "*"
                        }

                        Write-Verbose -Message "[i] Searching AD in OU '$SearchRootItem'"
                        Write-Verbose -Message ("[i] Filter: '{0}'" -f $ADGroupParams.Filter)

                        $GroupSearch = Get-ADGroup @ADGroupParams -ErrorAction Stop

                        if ($GroupSearch)
                        {
                            foreach ($GroupSearchItem in $GroupSearch)
                            {
                                $null = $GroupList.Add($GroupSearchItem.DistinguishedName)
                            }
                        }
                    }
                }
                'File'
                {
                    Write-Verbose -Message "[*] Get AD Groups from file"

                    foreach ($FileItem in $File)
                    {
                        Write-Verbose -Message "[*] Loading File: $FileItem"

                        $FileContent = Get-Content -Path $FileItem -ErrorAction Stop

                        if ($FileContent)
                        {
                            $null = $GroupList.Add($FileContent)
                        }
                    }
                }
                'Group'
                {
                    Write-Verbose -Message "[*] Get AD Groups from 'Group' parameter'"

                    foreach ($GroupItem in $Group)
                    {
                        $null = $GroupList.Add($GroupItem)
                    }
                }
            }

            # Process every group prioided by any ParameterSet
            foreach ($GroupItem in $GroupList)
            {
                $Changes, $GroupName, $ChangeHistoryList, $MemberList = $null

                try
                {
                    Write-Verbose -Message "[*] Processing $GroupItem"

                    # Splatting for the AD Group Request
                    $GroupSplatting = @{}
                    $GroupSplatting.Identity = $GroupItem

                    if ($PSBoundParameters['Server'])
                    {
                        $GroupSplatting.Server = $Server
                    }

                    # Look for Group
                    $GroupName = Get-ADGroup @GroupSplatting -Properties * -ErrorAction Continue -ErrorVariable ErrorProcessGetADGroup
                    #Write-Verbose -Message "    [*] Extracting Domain Name from $($GroupName.CanonicalName)"
                    $DomainName = ($GroupName.CanonicalName -split '/')[0]
                    $RealGroupName = $GroupName.Name

                    if ($GroupName)
                    {
                        $GroupMemberSplatting = @{}
                        $GroupMemberSplatting.Identity = $GroupName

                        # Get GroupName Membership
                        Write-Verbose -Message "    [*] Querying AD Group Membership"

                        # Add the Server if specified
                        if ($PSBoundParameters['Server'])
                        {
                            $GroupMemberSplatting.Server = $Server
                        }

                        if ($PSBoundParameters['Recursive'])
                        {
                            $GroupMemberSplatting.Recursive = $true
                        }

                        $ADGroupMembers = Get-ADGroupMember @GroupMemberSplatting -ErrorAction Stop -ErrorVariable ErrorProcessGetADGroupMember

                        $Members = foreach ($ADGroupMemberItem in $ADGroupMembers)
                        {   switch ($ADGroupMemberItem.objectClass) {
                                'computer' { Get-ADComputer -Identity $ADGroupMemberItem -ErrorAction SilentlyContinue }
                                'group' { Get-ADGroup -Identity $ADGroupMemberItem -ErrorAction SilentlyContinue }
                                'user' { Get-ADUser -Identity $ADGroupMemberItem -ErrorAction SilentlyContinue }
                                Default {}
                            }
                        }

                        if (-not ($Members))
                        {
                            # no members, add some info in $members to avoid the $null
                            # if the value is $null the compare-object won't work

                            Write-Verbose -Message "    [-] Group is empty"

                            $Members = [PSCustomObject]@{
                                Name           = 'No Member'
                                SamAccountName = 'No Member'
                            }
                        }

                        #$Members | Format-Table -AutoSize -Property DateTime, State, DisplayName, Name, SamAccountName, DN, DistinguishedName

                        # Current Group Membership File
                        # if the file doesn't exist, assume we don't have a record to refer to
                        $CurrentGroupMembershipFilePath = Join-Path -Path $ScriptPathCurrent -ChildPath ("{0}_{1}-membership.csv" -f $DomainName, $RealGroupName)

                        if (-not (Test-Path -Path $CurrentGroupMembershipFilePath))
                        {
                            Write-Verbose -Message "    [i] The file does not exist: $CurrentGroupMembershipFilePath"

                            $Members | Export-Csv -Path $CurrentGroupMembershipFilePath -NoTypeInformation -Encoding Unicode

                            Write-Verbose -Message "    [+] Exported current group membership into file '$CurrentGroupMembershipFilePath'"
                        }
                        else
                        {
                            Write-Verbose -Message "    [i] The file exists: $CurrentGroupMembershipFilePath"
                        }

                        $ImportCSV = Import-Csv -Path $CurrentGroupMembershipFilePath -ErrorAction Stop -ErrorVariable ErrorProcessImportCSV

                        Write-Verbose -Message "    [*] Comparing Current and Before"

                        $CompareResult = Compare-Object -ReferenceObject $Members -DifferenceObject $ImportCSV -Property SamAccountName -PassThru -ErrorAction Stop -ErrorVariable ErrorProcessCompareObject

                        $Changes = $CompareResult | Where-Object { $_.SamAccountName -ne 'No Member' } | ForEach-Object {
                            $DateTime = Get-Date -Format $ChangesDateTimeFormat
                            $State    = switch ($PSItem.SideIndicator)
                            {
                                '=>' { 'Removed' }
                                '<=' { 'Added' }
                            }

                            $PSItem | Add-Member -Name 'DateTime' -Value $DateTime -MemberType NoteProperty -Force
                            $PSItem | Add-Member -Name 'State' -Value $State -MemberType NoteProperty -Force -PassThru
                        }

                        if ($Changes -or $AlwaysReport)
                        {
                            # Found Changes or report anyway
                            if ($Changes)
                            {
                                Write-Verbose -Message "    [i] Found group membership changes"
                            }
                            else
                            {
                                Write-Verbose -Message "    [i] Found no group membership changes"
                            }

                            # Get the
                            $ChangeList = New-Object System.Collections.Generic.List[Object]

                            foreach ($ChangeItem in $Changes)
                            {
                                $ListItem = [ordered]@{}

                                foreach ($PropertyName in $MembershipChangeTableProperty)
                                {
                                    switch ($PropertyName)
                                    {
                                        'DN'
                                        {
                                            # woraround if the column name in CSV is 'DN'
                                            if ($ChangeItem.$PropertyName -ne '')
                                            {
                                                $PropertyName  = 'DistinguishedName'
                                                $PropertyValue = $ChangeItem.$PropertyName
                                            }
                                        }
                                        'DateTime'
                                        {
                                            $PropertyValue = Convert-DateTimeString -String $ChangeItem.DateTime -InputFormat $KnownInputFormat -OutputFormat $ReportChangesDateTimeFormat
                                        }
                                        Default
                                        {
                                            $PropertyValue = $ChangeItem.$PropertyName
                                        }
                                    }

                                    $ListItem.Add($PropertyName, $PropertyValue)
                                }

                                $null = $ChangeList.Add([PSCustomObject]$ListItem)
                            }

                            Write-Verbose -Message "    [+] Created Membership Change List"

                            # Get the Changes History from CSV file
                            $ChangeHistoryFile = Get-ChildItem -Path $ScriptPathHistory\$($DomainName)_$($RealGroupName)-ChangeHistory.csv -ErrorAction SilentlyContinue

                            if ($ChangeHistoryFile)
                            {
                                $ChangeHistoryFilePath = $ChangeHistoryFile.FullName

                                Write-Verbose -Message "    [i] Found Change History file '$ChangeHistoryFilePath'"

                                $ChangeHistoryList = New-Object System.Collections.Generic.List[Object]

                                $ImportCsvFile = Import-Csv -Path $ChangeHistoryFilePath -ErrorAction Stop -ErrorVariable ErrorProcessImportCSVChangeHistory

                                foreach ($CsvFileLine in $ImportCsvFile)
                                {
                                    $ListItem = [ordered]@{}

                                    foreach ($PropertyName in $ChangeHistoryTableProperty)
                                    {
                                        switch ($PropertyName)
                                        {
                                            'DN'
                                            {
                                                # woraround if the column name in CSV is 'DN'
                                                if ($CsvFileLine.$PropertyName -ne '')
                                                {
                                                    $PropertyName  = 'DistinguishedName'
                                                    $PropertyValue = $CsvFileLine.$PropertyName
                                                }
                                            }
                                            'DateTime'
                                            {
                                                $PropertyValue = Convert-DateTimeString -String $CsvFileLine.DateTime -InputFormat $KnownInputFormat -OutputFormat $ReportChangesDateTimeFormat
                                            }
                                            Default
                                            {
                                                $PropertyValue = $CsvFileLine.$PropertyName
                                            }
                                        }

                                        $ListItem.Add($PropertyName, $PropertyValue)
                                    }

                                    $null = $ChangeHistoryList.Add([PSCustomObject]$ListItem)
                                }

                                Write-Verbose -Message "    [+] Imported Change History from CSV"
                            }

                            if ($Changes)
                            {
                                $CsvPath = Join-Path -Path $ScriptPathHistory -ChildPath "$($DomainName)_$($RealGroupName)-ChangeHistory.csv"

                                if (-not (Test-Path -Path $CsvPath))
                                {
                                    $Changes | Select-Object -Property $ChangeHistoryCsvProperty |
                                        Export-Csv -Path $CsvPath -NoTypeInformation -Encoding Unicode -ErrorAction Stop
                                }
                                else
                                {
                                    $Changes | Select-Object -Property $ChangeHistoryCsvProperty |
                                        Export-Csv -Path $CsvPath -NoTypeInformation -Append -Encoding Unicode -ErrorAction Stop
                                }

                                Write-Verbose -Message "    [+] Saved changes to the group's ChangeHistory file"
                            }

                            if ($IncludeMembers)
                            {
                                $MemberList = New-Object System.Collections.Generic.List[Object]

                                foreach ($MemberItem in $Members)
                                {
                                    $ListItem = [ordered]@{}

                                    foreach ($PropertyName in $MembersTableProperty)
                                    {
                                        $PropertyValue = $MemberItem.$PropertyName
                                        $ListItem.Add($PropertyName, $PropertyValue)
                                    }

                                    $null = $MemberList.Add([PSCustomObject]$ListItem)
                                }

                                Write-Verbose -Message "    [+] Created Member List"
                            }

                            #Region Email
                            Write-Verbose -Message "    [*] Preparing the notification email"

                            # Preparing the body of the email
                            $body = "<h2>Group: $RealGroupName</h2>"
                            $body += "<p style=`"background-color:white; font-family:Calibri; font-size:10pt`">"
                            $body += "<b>SamAccountName:</b> $($GroupName.SamAccountName)<br>"
                            $body += "<b>Description:</b> $($GroupName.Description)<br>"
                            $body += "<b>DistinguishedName:</b> $($GroupName.DistinguishedName)<br>"
                            $body += "<b>CanonicalName:</b> $($GroupName.CanonicalName)<br>"
                            $body += "<b>SID:</b> $($GroupName.Sid.value)<br>"
                            $body += "<b>Scope/Type:</b> $($GroupName.GroupScope) / $($GroupName.GroupType)<br>"
                            $body += "<b>gidNumber:</b> $($GroupName.gidNumber)<br>"
                            $body += "</p>"

                            if ($ChangeList)
                            {
                                $Body += "<h3>Membership Change</h3>"
                                $Body += "<i>The membership of this group changed. See the following Added or Removed members.</i>"
                                $Body += $ChangeList |  Sort-Object -Property DateTime -Descending | ConvertTo-Html -Head $Head | Out-String
                                #$Body += "<br><br><br>"

                                $MailPriority = 'High'
                            }
                            else
                            {
                                $Body += "<h3>Membership Change</h3>"
                                $Body += "<i>No changes.</i>"

                                $MailPriority = 'Normal'
                            }

                            if ($ChangeHistoryList)
                            {
                                $Body += "<h3>Change History</h3>"
                                $Body += "<i>List of previous changes on this group observed by the script</i>"
                                $Body += $ChangeHistoryList | Sort-Object -Property DateTime -Descending | ConvertTo-Html -Fragment -PreContent $Head2 | Out-String
                            }

                            if ($MemberList)
                            {
                                $Body += "<h3>Members</h3>"
                                $Body += "<i>List of all members</i>"
                                $Body += $MemberList | Sort-Object -Property SamAccountName -Descending | ConvertTo-Html -Fragment -PreContent $Head2 | Out-String
                            }

                            $Body = $Body -replace "Added", "<font color=`"green`"><b>Added</b></font>"
                            $Body = $Body -replace "Removed", "<font color=`"red`"><b>Removed</b></font>"
                            $Body += $Report

                            $EmailSubject = "$EmailSubjectPrefix $($GroupName.SamAccountName) membership change detected"

                            if ($OneReport)
                            {
                                if ($OneReportMailPriority -eq 'High')
                                {
                                    # Priority is already (due to a change in another group) set to 'High', so no update needed
                                }
                                else
                                {
                                    $OneReportMailPriority = $MailPriority
                                }
                            }
                            else
                            {
                                # only send e-mail for each group if 'OneReport' is not specified

                                $paramMailMessage = @{
                                    To         = $EmailTo
                                    From       = $EmailFrom
                                    Subject    = $EmailSubject
                                    Body       = $Body
                                    Priority   = $MailPriority
                                    SmtpServer = $EmailServer
                                    Port       = $EmailPort
                                    #UseSsl     = $true
                                    BodyAsHtml  = $true
                                }

                                if ($EmailCredential -ne [System.Management.Automation.PSCredential]::Empty) {
                                    $paramMailMessage['Credential'] = $EmailCredential
                                }

                                Send-MailMessage @paramMailMessage

                                Write-Verbose -Message "    [+] Sent Email"
                            }
                            #EndRegion

                            # GroupName Membership export to CSV
                            $Members | Export-Csv -Path $CurrentGroupMembershipFilePath -NoTypeInformation -Encoding Unicode -ErrorAction Stop
                            Write-Verbose -Message "    [*] Exported current group membership to file '$CurrentGroupMembershipFilePath'"


                            # Define HTML File Name
                            #$HTMLFileName = "$($DomainName)_$($RealGroupName)-$(Get-Date -Format $FileNameDateTimeFormat).html"
                            $HTMLFileName = "{0}_{1}-{2}.html" -f $DomainName, $RealGroupName, (Get-Date -Format $FileNameDateTimeFormat)

                            if ($PSBoundParameters['SaveAsHTML'])
                            {
                                # Save HTML File
                                $Body | Out-File -FilePath (Join-Path -Path $ScriptPathHTML -ChildPath $HTMLFileName) -ErrorAction Stop

                                Write-Verbose -Message "    [+] Saved as HTML file"
                            }

                            if ($OneReport)
                            {
                                # Save HTML File
                                $Body | Out-File -FilePath (Join-Path -Path $ScriptPathOneReport -ChildPath $HTMLFileName) -ErrorAction Stop

                                Write-Verbose -Message "    [+] Saved as HTML file for OneReport"
                            }
                        }
                        else
                        {
                            Write-Verbose -Message "    [i] Found no group membership changes"
                        }
                    }
                    else
                    {
                        Write-Verbose -Message "    [*] Group can't be found"
                    }
                }
                catch
                {
                    Write-Warning -Message "    [!] Something went wrong"

                    # Active Directory Module Errors
                    if ($ErrorProcessGetADGroup) { Write-Warning -Message "    [!] Error querying the group $GroupItem in Active Directory"; }
                    if ($ErrorProcessGetADGroupMember) { Write-Warning -Message "    [!] Error querying the group $GroupItem members in Active Directory"; }

                    # Import CSV Errors
                    if ($ErrorProcessImportCSV) { Write-Warning -Message "    [!] Error importing $CurrentGroupMembershipFilePath"; }
                    if ($ErrorProcessCompareObject) { Write-Warning -Message "    [!] Error when comparing"; }
                    if ($ErrorProcessImportCSVChangeHistory) { Write-Warning -Message "    [!] Error importing $ChangeHistoryFilePath"; }

                    Write-Warning -Message $_.Exception.Message
                }
            }

            if ($OneReport)
            {
                #[string[]]$Attachments = @()
                $Attachments = foreach ($File in (Get-ChildItem $ScriptPathOneReport))
                {
                    #$Attachments.Add($File.FullName)
                    $File.FullName
                }

                $EmailSubject = "$EmailSubjectPrefix Membership Report"

                $paramMailMessage = @{
                    To          = $EmailTo
                    From        = $EmailFrom
                    Subject     = $EmailSubject
                    Body        = "<h2>See Report in Attachment</h2>"
                    Priority    = $OneReportMailPriority
                    SmtpServer  = $EmailServer
                    Port        = $EmailPort
                    #UseSsl     = $true
                    Attachments = $Attachments
                    BodyAsHtml  = $true
                }

                if ($EmailCredential -ne [System.Management.Automation.PSCredential]::Empty) {
                    $paramMailMessage['Credential'] = $EmailCredential
                }

                Send-MailMessage @paramMailMessage

                foreach ($a in $Attachments)
                {
                    $a.Dispose()
                }

                Get-ChildItem $ScriptPathOneReport | Remove-Item -Force -Confirm:$false
                Write-Verbose -Message "    [+] Sent OneReport Email"
            }
        }
        catch
        {
            Write-Warning -Message "    [*] Something went wrong"
            throw $_.Exception.Message
        }
    }
    End
    {
        $StopWatch.Stop()

        Write-Verbose -Message "[$ScriptName][End]"
        Write-Verbose -Message ("    [i] Elapsed time: {0:d2}:{1:d2}:{2:d2}" -f $StopWatch.Elapsed.Hours, $StopWatch.Elapsed.Minutes, $StopWatch.Elapsed.Seconds)
        Write-Verbose -Message "    [i] Script Completed"
    }
}
