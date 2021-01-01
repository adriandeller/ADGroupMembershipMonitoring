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
        Default is UTF8.
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
    .PARAMETER ExcludeChanges
        Specify if you want to exclude changes in the Report.
    .PARAMETER ExcludeHistory
        Specify if you want to exclude the change history in the Report.
    .PARAMETER ExcludeSummary
        Specify if you want to exclude the summary at the top of the Report.
    .PARAMETER AlwaysReport
        Specify if you want to generate a HTML Report and send an Email each time.
    .PARAMETER AlwaysExport
        Specify if you want to generate a HTML Report each time.
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
        Updated:     2021-01-02
        Version:     0.3.0
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
        $EmailEncoding = 'UTF8',

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
        $ExcludeChanges,

        [Parameter()]
        [Switch]
        $ExcludeHistory,

        [Parameter()]
        [Switch]
        $ExcludeSummary,

        [Parameter()]
        [Switch]
        $AlwaysExport,

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

            if ($PSBoundParameters['EmailSubjectPrefix'])
            {
                $EmailSubjectPrefix = $PSBoundParameters['EmailSubjectPrefix']
            }
            else
            {
                $EmailSubjectPrefix = $Configuration.Email.EmailSubjectPrefix
            }

            # CSV columns
            $ChangeHistoryCsvProperty = 'DateTime', 'State', 'DisplayName', 'Name', 'SamAccountName', 'ObjectClass', 'DistinguishedName'
            #$ChangeHistoryCsvProperty = $Configuration.CSV.ChangeHistoryProperty
            $MembershipCsvProperty    = 'Name', 'SamAccountName', 'ObjectClass', 'Mail', 'UserPrincipalName'
            #$MembershipCsvProperty = $Configuration.CSV.MembershipProperty

            # Report table columns
            $GroupSummaryTableProperty = 'SamAccountName', 'Description', 'DistinguishedName', 'CanonicalName', 'SID', 'GroupScope', 'GroupCategory', 'gidNumber'
            $MembershipChangeTableProperty = 'DateTime', 'State', 'Name', 'SamAccountName', 'ObjectClass', 'DistinguishedName'
            $ChangeHistoryTableProperty = 'DateTime', 'State', 'Name', 'SamAccountName', 'ObjectClass'
            $MembersTableProperty = 'Name', 'SamAccountName', 'ObjectClass', 'Mail'

            $SpecialProperty = 'DateTime', 'State'

            if ($PSBoundParameters['ExtendedProperty'])
            {
                $MembersTableProperty += 'Enabled', 'PasswordExpired', 'DistinguishedName'
            }

            # Create array with all required properties/attributes for AD object queries
            $ADObjectPropertyArray = $ChangeHistoryCsvProperty + $MembershipCsvProperty + $GroupSummaryTableProperty + $MembershipChangeTableProperty + $ChangeHistoryTableProperty + $MembersTableProperty |
                Select-Object -Unique -ExcludeProperty $SpecialProperty

            # Filter properties for valid Computer object attributes
            $ADComputerAttributes = Get-ADObjectAttributes -ClassName 'computer' -Verbose:$false
            $ADComputerProperty = $ADObjectPropertyArray | Where-Object { $PSItem -in $ADComputerAttributes }

            # Filter properties for valid User object attributes
            $ADUserAttributes = Get-ADObjectAttributes -ClassName 'user' -Verbose:$false
            $ADUserProperty = $ADObjectPropertyArray | Where-Object { $PSItem -in $ADUserAttributes }

            # Filter properties for valid Group object attributes
            $ADGroupAttributes = Get-ADObjectAttributes -ClassName 'group' -Verbose:$false
            $ADGroupProperty = $ADObjectPropertyArray | Where-Object { $PSItem -in $ADGroupAttributes }

            $PSDefaultParameterValues = @{
                'Convert-DateTimeString:Verbose' = $false
                'Get-ADObjectDetails:Verbose' = $false
                'Get-ADComputer:Property' = $ADComputerProperty
                'Get-ADUser:Property' = $ADUserProperty
                'Get-ADGroup:Property' = $ADGroupProperty
            }
            #

            # Set the Date and Time formats
            $ChangesDateTimeFormat = 's'                         # format for export to CSV files
            $ReporCreatedDateTimeFormat = 'dd.MM.yyyy HH:mm:ss'  # format for report creation date/time information
            $ReportChangesDateTimeFormat = 'yyyy-MM-dd HH:mm:ss' # format for DateTime property in HTML reports
            $FileNameDateTimeFormat = 'yyyyMMdd_HHmmss'          # format for DateTime information in CSV file names

            # list of known DateTime formats used in legacy CSV files
            $KnownInputFormat = 'yyyyMMdd-HH:mm:ss', $ChangesDateTimeFormat

            # HTML blocks for Report
            $PostContent = "<p style=`"font-size:10pt`">" +
            "<strong>Report created:</strong> $(Get-Date -Format $ReporCreatedDateTimeFormat)<br>" +
            "<strong>Account:</strong> $env:USERDOMAIN\$($env:USERNAME) on $($env:COMPUTERNAME)" +
            "</p>"

            $Head = "<style>" +
            "body {background-color:white; font-family:Calibri; font-size:11pt;}" +
            "p {font-size:10pt; margin-bottom:10;}" +
            "h2 {font-family:Calibri;}" +
            "h3 {font-family:Calibri; margin-top:40px;}" +
            "table {border-width:1px; border-style:solid; border-color:black; border-collapse:collapse;}" +
            "th {border-width:1px; padding:3px; border-style:solid; border-color:black; background-color:#6495ED;}" +
            "td {border-width:1px; padding-right:2px; padding-left:2px; padding-top:0px; padding-bottom:0px; border-style:solid; border-color:black; background-color:white;}" +
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
        Write-Verbose -Message "[$ScriptName][Process]"

        try
        {
            #Region Input handling based on ParameterSetName
            switch ($PSCmdlet.ParameterSetName)
            {
                'ADFilter'
                {
                    Write-Verbose -Message "[*] Using ParameterSet 'ADFilter'"

                    $SearchRootItemNumber = 1

                    foreach ($SearchRootItem in $SearchRoot)
                    {
                        # ADGroup Splatting
                        $ADGroupParams = @{ }

                        Write-Verbose -Message "    [i] [$SearchRootItemNumber] SearchRoot: $SearchRootItem"
                        $ADGroupParams.SearchBase = $SearchRootItem

                        # Server parameter specified
                        if ($PSBoundParameters['Server'])
                        {
                            Write-Verbose -Message "    [i] [$SearchRootItemNumber] Server: $Server"
                            $ADGroupParams.Server = $Server
                        }

                        # SearchScope parameter specified
                        if ($PSBoundParameters['SearchScope'])
                        {
                            Write-Verbose -Message "    [i] [$SearchRootItemNumber] SearchScope: $SearchScope"
                            $ADGroupParams.SearchScope = $SearchScope
                        }

                        # GroupScope parameter specified
                        if ($PSBoundParameters['GroupScope'])
                        {
                            Write-Verbose -Message "    [i] [$SearchRootItemNumber] GroupScope: $GroupScope"
                            $ADGroupParams.Filter = "GroupScope -eq `'$GroupScope`'"
                        }

                        # GroupType parameter specified
                        if ($PSBoundParameters['GroupType'])
                        {
                            Write-Verbose -Message "    [i] [$SearchRootItemNumber] GroupType: $GroupType"

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
                            Write-Verbose -Message "    [i] [$SearchRootItemNumber] GroupFilter: $GroupFilter"

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

                        Write-Verbose -Message ("    [i] [$SearchRootItemNumber] Filter: '{0}'" -f $ADGroupParams.Filter)

                        $GroupSearch = Get-ADGroup @ADGroupParams -ErrorAction Stop

                        if ($GroupSearch)
                        {
                            foreach ($GroupSearchItem in $GroupSearch)
                            {
                                #$null = $GroupList.Add($GroupSearchItem.DistinguishedName)
                                $null = $GroupList.Add($GroupSearchItem)
                            }
                        }

                        $SearchRootItemNumber++
                    }
                }
                'File'
                {
                    Write-Verbose -Message "[*] Using ParameterSet 'File'"

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
                    Write-Verbose -Message "[*] Using ParameterSet 'Group'"

                    foreach ($GroupItem in $Group)
                    {
                        $null = $GroupList.Add($GroupItem)
                    }
                }
            }
            #EndRegion

            #Region Process groups
            # Process every group proided by any ParameterSet
            foreach ($GroupListItem in $GroupList)
            {
                $Changes, $ADGroup, $ChangeList, $ChangeHistoryList, $MemberList = $null
                $PreContent, $HtmlGroupSummary, $HtmlChangeList, $HtmlChangeHistoryList, $HtmlMemberList, $PostContent = $null

                try
                {
                    Write-Verbose -Message "[*] Processing $GroupListItem"

                    if ($GroupListItem -is [Microsoft.ActiveDirectory.Management.ADGroup])
                    {
                        # Input is already an ADGroup object

                        $ADGroup = $GroupListItem

                        Write-Verbose -Message "    [i] Input is an ADGroup object"
                    }
                    else
                    {
                        # Input is a DistinguishedName

                        $GroupSplatting = @{}
                        $GroupSplatting.Identity = $GroupListItem

                        if ($PSBoundParameters['Server'])
                        {
                            $GroupSplatting.Server = $Server
                        }
                        # Search for Group
                        $ADGroup = Get-ADGroup @GroupSplatting -ErrorAction Stop -ErrorVariable ErrorProcessGetADGroup

                        Write-Verbose -Message "    [+] Searched for AD Group"
                    }

                    # Get Domain (FQDN) of the AD Group
                    $DomainName = (Get-ADDomain -Identity $ADGroup.SID.AccountDomainSid.value -ErrorAction Stop).DNSRoot
                    $ADGroupName = $ADGroup.Name

                    #Region Get Group Membership
                    Write-Verbose -Message "    [*] Querying AD Group Membership"

                    $GroupMemberSplatting = @{}
                    $GroupMemberSplatting.Identity = $ADGroup

                    if ($PSBoundParameters['Server'])
                    {
                        $GroupMemberSplatting.Server = $Server
                    }

                    if ($PSBoundParameters['Recursive'])
                    {
                        $GroupMemberSplatting.Recursive = $true
                    }

                    $ADGroupMembers = Get-ADGroupMember @GroupMemberSplatting -ErrorAction Stop -ErrorVariable ErrorProcessGetADGroupMember

                    # For each group member, get AD object with needed properties
                    $Members = $ADGroupMembers | Get-ADObjectDetails -ErrorAction Stop

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
                    #EndRegion

                    #Region Find Changes

                    # Current Group Membership File
                    # if the file doesn't exist, assume we don't have a record to refer to
                    $CurrentGroupMembershipFilePath = Join-Path -Path $ScriptPathCurrent -ChildPath ("{0}_{1}-Membership.csv" -f $DomainName, $ADGroupName)

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

                    $MembersFromCsv = Import-Csv -Path $CurrentGroupMembershipFilePath -ErrorAction Stop -ErrorVariable ErrorProcessImportCSV

                    # For each member in Csv file, get AD object with needed properties, if exists
                    $MembersFromCsvDetails = $MembersFromCsv | Get-ADObjectDetails -ErrorAction Stop

                    Write-Verbose -Message "    [*] Comparing Current and Before"

                    $CompareResult = Compare-Object -ReferenceObject $Members -DifferenceObject $MembersFromCsvDetails -Property SamAccountName -PassThru -ErrorAction Stop -ErrorVariable ErrorProcessCompareObject

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
                    #EndRegion


                    if ($Changes -or $AlwaysReport -or $AlwaysExport)
                    {
                        # Found Changes or report/export anyway
                        if ($Changes)
                        {
                            Write-Verbose -Message "    [i] Found group membership changes"
                        }
                        else
                        {
                            Write-Verbose -Message "    [i] Found no group membership changes"
                        }

                        #Region Process Changes
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
                                        # Woraround if the column name in CSV is 'DN'
                                        if ($ChangeItem.$PropertyName -ne '')
                                        {
                                            $PropertyName  = 'DistinguishedName'
                                            $PropertyValue = $ChangeItem.$PropertyName
                                        }
                                    }
                                    'DateTime'
                                    {
                                        # Convert DateTime sting to custom format
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
                        #EndREgion

                        #Region Get the Changes History from CSV file
                        $ChangeHistoryFile = Get-ChildItem -Path $ScriptPathHistory\$($DomainName)_$($ADGroupName)-ChangeHistory.csv -ErrorAction SilentlyContinue

                        if ($ChangeHistoryFile)
                        {
                            $ChangeHistoryFilePath = $ChangeHistoryFile.FullName

                            Write-Verbose -Message "    [i] Found Change History file '$ChangeHistoryFilePath'"

                            $ChangeHistoryList = New-Object System.Collections.Generic.List[Object]

                            $ChangeHistoryFromCsv = Import-Csv -Path $ChangeHistoryFilePath -ErrorAction Stop -ErrorVariable ErrorProcessImportCSVChangeHistory

                            foreach ($ChangeHistoryItem in $ChangeHistoryFromCsv)
                            {
                                $ADObjectDetails = $null

                                try
                                {
                                    # For the member in Csv file, get the AD object with needed properties, if exists
                                    $ADObjectDetails = $ChangeHistoryItem | Get-ADObjectDetails -ErrorAction Stop
                                }
                                catch
                                {
                                    Write-Warning $_.Exception.Message
                                }

                                $ListItem = [ordered]@{}

                                foreach ($PropertyName in $ChangeHistoryTableProperty)
                                {
                                    $PropertyValue = $null

                                    switch ($PropertyName)
                                    {
                                        'DateTime'
                                        {
                                            # handling property for which the data is only available in a column in the CSV file
                                            $PropertyValue = Convert-DateTimeString -String $ChangeHistoryItem.DateTime -InputFormat $KnownInputFormat -OutputFormat $ReportChangesDateTimeFormat
                                            #Write-Verbose "$PropertyName : $PropertyValue"
                                        }
                                        'State'
                                        {
                                            # handling property for which the data is only available in a column in the CSV file
                                            $PropertyValue = $ChangeHistoryItem.$PropertyName
                                            #Write-Verbose "$PropertyName : $PropertyValue"
                                        }
                                        Default
                                        {
                                            if ($ADObjectDetails)
                                            {
                                                $PropertyValue = $ADObjectDetails.$PropertyName
                                                #Write-Verbose "$PropertyName : $PropertyValue"
                                            }
                                            else
                                            {
                                                if ($PropertyName -in ('SamAccountName', 'ObjectClass'))
                                                {
                                                    # handling exception where AD object is not found, return the identifier attribute 'SamAccountName' value from CSV file
                                                    $PropertyValue = $ChangeHistoryItem.$PropertyName
                                                    #Write-Verbose "$PropertyName : $PropertyValue (from CSV)"
                                                }
                                                else
                                                {
                                                    Write-Verbose "No data in CSV file for property '$PropertyName'"
                                                }
                                            }
                                        }
                                    }

                                    $ListItem.Add($PropertyName, $PropertyValue)
                                }

                                $null = $ChangeHistoryList.Add([PSCustomObject]$ListItem)
                            }

                            Write-Verbose -Message "    [+] Imported changes from ChangeHistory CSV file"
                        }
                        #EndRegion

                        #Region Export Changes to ChangeHistory CSV file
                        if ($Changes)
                        {
                            $ChangeHistoryFile = Join-Path -Path $ScriptPathHistory -ChildPath "$($DomainName)_$($ADGroupName)-ChangeHistory.csv"

                            if (Test-Path -Path $ChangeHistoryFile)
                            {
                                # ChangeHistory file exists and content is already imported to variable $ChangeHistoryFromCsv
                                # - compare CSV columns with properties to export
                                # - add missing columns if required
                                # - combine existing and new Changes and export to CSV
                                $CsvColumns = $ChangeHistoryFromCsv[0].psobject.Properties.Name
                                $MissingCsvColumns = $ChangeHistoryCsvProperty | Where-Object { $PSItem -notin $CsvColumns }
                                if ($MissingCsvColumns)
                                {
                                    foreach ($ChangeHistoryItem in $ChangeHistoryFromCsv)
                                    {
                                        foreach ($CsvColumn in $MissingCsvColumns) {
                                            $ChangeHistoryItem | Add-Member -MemberType NoteProperty -Name $CsvColumn -Value $null
                                        }
                                    }
                                }
                                $ChangesToExport = $ChangeHistoryFromCsv + $Changes
                                $ChangesToExport | Select-Object -Property $ChangeHistoryCsvProperty |
                                    Export-Csv -Path $ChangeHistoryFile -NoTypeInformation -Encoding Unicode -ErrorAction Stop

                                <#
                                $Changes | Select-Object -Property $ChangeHistoryCsvProperty |
                                    Export-Csv -Path $ChangeHistoryFile -NoTypeInformation -Append -Encoding Unicode -ErrorAction Stop
                                #>
                            }
                            else
                            {
                                $Changes | Select-Object -Property $ChangeHistoryCsvProperty |
                                    Export-Csv -Path $ChangeHistoryFile -NoTypeInformation -Encoding Unicode -ErrorAction Stop
                            }

                            Write-Verbose -Message "    [+] Exported changes to ChangeHistory CSV file '$ChangeHistoryFile'"
                        }
                        #EndRegion

                        #Region Members
                        if ($Changes)
                        {
                            # Export Current Group Membership to CSV
                            $Members | Select-Object -Property $MembershipCsvProperty |
                                Export-Csv -Path $CurrentGroupMembershipFilePath -NoTypeInformation -Encoding Unicode -ErrorAction Stop
                            Write-Verbose -Message "    [*] Exported current group membership to Membership CSV file '$CurrentGroupMembershipFilePath'"
                        }

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
                        #EndRegion

                        #Region Group Summary
                        $GroupSummary = [ordered]@{}

                        foreach ($PropertyName in $GroupSummaryTableProperty)
                        {
                            switch ($PropertyName)
                            {
                                'SID'
                                {
                                    $PropertyValue = $ADGroup.$PropertyName.value
                                    $GroupSummary.Add($PropertyName, $PropertyValue)
                                }
                                Default
                                {
                                    $PropertyValue = $ADGroup.$PropertyName
                                    $GroupSummary.Add($PropertyName, $PropertyValue)
                                }
                            }
                        }

                        $GroupSummary = [PSCustomObject]$GroupSummary
                        #EndRegion

                        #Region HTML Report / Email
                        Write-Verbose -Message "    [*] Preparing the HTML report and Email"

                        # Preparing the body of the email
                        $PreContent = "<h2>Group: $ADGroupName</h2>"

                        if (-not ($PSBoundParameters['ExcludeSummary']))
                        {
                            $HtmlGroupSummaryPreContent = "<h3>Summary</h3>"
                            $HtmlGroupSummary = $GroupSummary | ConvertTo-Html -PreContent $HtmlGroupSummaryPreContent -Fragment -As List | Out-String
                        }

                        if ($ChangeList -and (-not ($PSBoundParameters['ExcludeChanges'])))
                        {
                            $HtmlChangeListPreContent = "<h3>Membership Change</h3>"
                            $HtmlChangeListPreContent += "<p><i>The membership of this group changed. See the following Added or Removed members.</i></p>"
                            $HtmlChangeList = $ChangeList | Sort-Object -Property DateTime -Descending | ConvertTo-Html -PreContent $HtmlChangeListPreContent -Fragment | Out-String

                            $MailPriority = 'High'
                        }
                        else
                        {
                            $Body += "<h3>Membership Change</h3>"
                            $Body += "<p><i>No changes.</i></p>"

                            $MailPriority = 'Normal'
                        }

                        if ($ChangeHistoryList -and (-not ($PSBoundParameters['ExcludeHistory'])))
                        {
                            $HtmlChangeHistoryListPreContent = "<h3>Change History</h3>"
                            $HtmlChangeHistoryListPreContent += "<p><i>List of previous changes on this group observed by the script</i></p>"
                            $HtmlChangeHistoryList = $ChangeHistoryList | Sort-Object -Property DateTime -Descending | ConvertTo-Html -PreContent $HtmlChangeHistoryListPreContent -Fragment | Out-String

                            $HtmlChangeHistoryList = $HtmlChangeHistoryList -replace "Added", "<font color=`"green`"><b>Added</b></font>"
                            $HtmlChangeHistoryList = $HtmlChangeHistoryList -replace "Removed", "<font color=`"red`"><b>Removed</b></font>"
                        }

                        if ($MemberList -and $PSBoundParameters['IncludeMembers'])
                        {
                            $HtmlMemberListPreContent = "<h3>Members</h3>"
                            $HtmlMemberListPreContent += "<p><i>List of all members</i></p>"
                            $HtmlMemberList = $MemberList | Sort-Object -Property SamAccountName -Descending | ConvertTo-Html -PreContent $HtmlMemberListPreContent -Fragment | Out-String
                        }

                        $Body = ConvertTo-Html -Head $Head -Body $PreContent, $HtmlGroupSummary, $HtmlChangeList, $HtmlChangeHistoryList, $HtmlMemberList, $PostContent | Out-String

                        if ($PSBoundParameters['OneReport'])
                        {
                            # do not send e-mail for each group if 'OneReport' is specified
                            if ($OneReportMailPriority -eq 'High')
                            {
                                # Priority is already (due to a change in another group) set to 'High', so no update needed
                            }
                            else
                            {
                                $OneReportMailPriority = $MailPriority
                            }
                        }
                        elseif ($Changes -or $PSBoundParameters['AlwaysReport'])
                        {
                            # send e-mail for each group if Changes found or if 'AlwaysReport' is specified

                            if ($Changes)
                            {
                                $EmailSubject = "$EmailSubjectPrefix $ADGroupName - changed"
                            }
                            else
                            {
                                $EmailSubject = "$EmailSubjectPrefix $ADGroupName"
                            }

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

                            Send-MailMessage @paramMailMessage -ErrorAction Stop

                            Write-Verbose -Message "    [+] Sent Email"
                        }
                        #EndRegion

                        #Region Export HTML report to file
                        $HTMLFileName = "{0}_{1}-{2}.html" -f $DomainName, $ADGroupName, (Get-Date -Format $FileNameDateTimeFormat)

                        if ($PSBoundParameters['SaveAsHTML'] -or $PSBoundParameters['AlwaysExport'])
                        {
                            # Save HTML File
                            $Body | Out-File -FilePath (Join-Path -Path $ScriptPathHTML -ChildPath $HTMLFileName) -ErrorAction Stop

                            Write-Verbose -Message "    [+] Saved as HTML file"
                        }

                        if ($PSBoundParameters['OneReport'])
                        {
                            # Save HTML File
                            $Body | Out-File -FilePath (Join-Path -Path $ScriptPathOneReport -ChildPath $HTMLFileName) -ErrorAction Stop

                            Write-Verbose -Message "    [+] Saved as HTML file for OneReport"
                        }
                        #EndRegion
                    }
                    else
                    {
                        Write-Verbose -Message "    [i] Found no group membership changes"
                    }
                }
                catch
                {
                    Write-Warning -Message "    [!] Something went wrong"

                    # Active Directory Module Errors
                    if ($ErrorProcessGetADGroup) { Write-Warning -Message "    [!] Error querying the group $GroupListItem in Active Directory"; }
                    if ($ErrorProcessGetADGroupMember) { Write-Warning -Message "    [!] Error querying the group $GroupListItem members in Active Directory"; }

                    # Import CSV Errors
                    if ($ErrorProcessImportCSV) { Write-Warning -Message "    [!] Error importing $CurrentGroupMembershipFilePath"; }
                    if ($ErrorProcessCompareObject) { Write-Warning -Message "    [!] Error when comparing"; }
                    if ($ErrorProcessImportCSVChangeHistory) { Write-Warning -Message "    [!] Error importing $ChangeHistoryFilePath"; }

                    Write-Warning -Message $_.Exception.Message
                }
            }
            #EndRegion

            #Region Send OneReport Email
            if ($PSBoundParameters['OneReport'] -and ($Changes -or $PSBoundParameters['AlwaysReport']))
            {
                # send OneReport e-mail if 'OneReport' is specified and
                # if Changes found or if 'AlwaysReport' is specified

                $Attachments = foreach ($OneReportFile in (Get-ChildItem -File -Path $ScriptPathOneReport))
                {
                    $OneReportFile.FullName
                }

                if ($Changes)
                {
                    $EmailSubject = "$EmailSubjectPrefix - Changes detected"
                }
                else
                {
                    $EmailSubject = "$EmailSubjectPrefix"
                }

                $paramMailMessage = @{
                    To          = $EmailTo
                    From        = $EmailFrom
                    Subject     = $EmailSubject
                    Body        = "<b>See the report(s) in the attachment</b>"
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

                Send-MailMessage @paramMailMessage -ErrorAction Stop

                Get-ChildItem $ScriptPathOneReport | Remove-Item -Force -Confirm:$false
                Write-Verbose -Message "    [+] Sent OneReport Email"
            }
            #EndRegion
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
