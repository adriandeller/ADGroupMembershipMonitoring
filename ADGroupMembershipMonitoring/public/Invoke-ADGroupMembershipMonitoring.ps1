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
        By default, search for group members with direct membership,
        Specify this switch and group members with indirect membership through group nesting will also be searched
    .PARAMETER SearchRoot
        Specify the DN, GUID or canonical name of the domain or container to search.
        By default, the script searches the entire sub-tree of which SearchRoot is the topmost object (sub-tree search).
        This default behavior can be altered by using the SearchScope parameter.
    .PARAMETER SearchScope
        Specify one of these parameter values
            'Base' Limits the search to the base (SearchRoot) object.
                The result contains a maximum of one object.
            'OneLevel' Searches the immediate child objects of the base (SearchRoot) object, excluding the base object.
            'Subtree'   Searches the whole sub-tree, including the base (SearchRoot) object and all its child objects.
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
    .PARAMETER LDAPFilter
        Specify the LDAP filter you want to use to find groups
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
    .PARAMETER SendEmail
        Specify if you want to send an Email including the HTML Report.
    .PARAMETER SaveReport
        Specify if you want to save the HTML Report.
        It will be saved under the "HTML" directory.
    .PARAMETER IncludeMembers
        Specify if you want to include all members in the Report.
    .PARAMETER ExcludeChanges
        Specify if you want to exclude changes in the Report.
    .PARAMETER ExcludeHistory
        Specify if you want to exclude the change history in the Report.
    .PARAMETER ExcludeSummary
        Specify if you want to exclude the summary at the top of the Report.
    .PARAMETER ForceAction
        Specify if you want to send an Email (with the HTML Report) and save a HTML Report each time.
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

        [Parameter(ParameterSetName = 'ADFilter', HelpMessage = 'You can specify multiple Active Directory OU')]
        [Parameter(ParameterSetName = 'LDAPFilter', HelpMessage = 'You can specify multiple Active Directory OU')]
        [Alias('SearchBase')]
        [string[]]
        $SearchRoot,

        [Parameter(ParameterSetName = 'ADFilter')]
        [Parameter(ParameterSetName = 'LDAPFilter')]
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

        [Parameter(ParameterSetName = 'ADFilter', Mandatory)]
        [string]
        $GroupFilter,

        [Parameter(ParameterSetName = 'LDAPFilter', Mandatory)]
        [string]
        $LDAPFilter,

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
        [Alias('Domain', 'DomainController', 'Service')]
        [string]
        $Server,

        [Parameter(HelpMessage = 'You must specify the sender E-Mail Address')]
        [ValidatePattern("[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?")]
        [string]
        $EmailFrom,

        [Parameter(HelpMessage = 'You must specify the recipient(s) E-Mail Address')]
        [ValidatePattern("[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?")]
        [string[]]
        $EmailTo,

        [Parameter(HelpMessage = 'You must specify the Mail Server to use (FQDN or ip address)')]
        [string]
        $EmailServer,

        [Parameter(HelpMessage = 'You can specify an alternate port on the (SMTP) Mail Server')]
        [string]
        $EmailPort = '25',

        [Parameter(HelpMessage = 'You can specify a prefix for the E-Mail subject')]
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
        [switch]
        $SendEmail,

        [Parameter()]
        [Switch]
        $SaveReport,

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
        $ForceAction,

        [Parameter()]
        [Switch]
        $OneReport,

        [Parameter()]
        [string[]]
        $MembersProperty,

        [Parameter(Mandatory, HelpMessage = 'You must specify a path for data storage')]
        [Alias('OutputPath', 'FolderPath')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Path
    )

    Begin
    {
        #Region Parameter validation
        if ($PSBoundParameters.ContainsKey('SendEmail'))
        {
            if(-not ($PSBoundParameters.ContainsKey('EmailFrom') -and $PSBoundParameters.ContainsKey('EmailTo') -and $PSBoundParameters.ContainsKey('EmailServer')))
            {
                throw "The 'EmailFrom', 'EmailTo' and 'EmailServer' parameters are required when using SendEmail"
            }
        }

        if ($PSBoundParameters.ContainsKey('OneReport'))
        {
            if (-not $PSBoundParameters.ContainsKey('SendEmail'))
            {
                throw "The 'SendEmail' parameter is required when using OneReport"
            }
        }
        #EndRegion

        $ScriptName = $MyInvocation.MyCommand.Name

        Write-Verbose -Message "[$ScriptName][Begin]"

        $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()

        $GroupList = New-Object System.Collections.Generic.List[Object]

        try
        {
            #Region Import Modules
            $RequiredModules = 'ActiveDirectory', 'PSFramework'

            foreach ($ModuleItem in $RequiredModules)
            {
                if (Get-Module -ListAvailable -Name $ModuleItem -Verbose:$false)
                {
                    Import-Module -Name $ModuleItem -Verbose:$false
                    Write-Verbose -Message "    [+] Imported module '$ModuleItem'"
                }
                else
                {
                    Write-Warning "    [-] Module '$ModuleItem' does not exist"
                    return
                }
            }
            #EndRegion

            #Region Configuration
            # Using PSFramework for configuration
            if (-not (Get-PSFConfig -Module $ModuleName))
            {
                throw 'Unable to get configuration data using PSFramework'
            }

            if ($PSBoundParameters['EmailSubjectPrefix'])
            {
                $EmailSubjectPrefix = $PSBoundParameters['EmailSubjectPrefix']
            }
            else
            {
                $EmailSubjectPrefix = Get-PSFConfigValue -FullName  "$ModuleName.Email.EmailSubjectPrefix"
            }

            # CSV columns
            $ChangeHistoryCsvProperty = Get-PSFConfigValue -FullName  "$ModuleName.CSV.ChangeHistoryProperty"
            $MembershipCsvProperty = Get-PSFConfigValue -FullName  "$ModuleName.CSV.MembershipProperty"

            # Report table columns
            $GroupSummaryTableProperty = 'SamAccountName', 'Description', 'DistinguishedName', 'CanonicalName', 'SID', 'GroupScope', 'GroupCategory', 'gidNumber'
            $MembershipChangeTableProperty = 'DateTime', 'State', 'Name', 'SamAccountName', 'ObjectClass'
            $ChangeHistoryTableProperty = 'DateTime', 'State', 'Name', 'SamAccountName', 'ObjectClass'
            $MembersTableProperty = 'Name', 'SamAccountName', 'ObjectClass'

            $SpecialProperty = 'DateTime', 'State'

            if ($PSBoundParameters['Recursive'])
            {
                $MembersTableProperty += 'Nested'
            }

            if ($PSBoundParameters['MembersProperty'])
            {
                $MembersTableProperty += $PSBoundParameters['MembersProperty']
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
            $ReporModuleName     = $Script:ModuleName
            $ReportModuleVersion = (Get-Module -Name $ReporModuleName).Version
            $PostContent = "<h3>Report</h3>"
            $PostContent += "<p>" +
            "<strong>Created:</strong> $(Get-Date -Format $ReporCreatedDateTimeFormat)<br>" +
            "<strong>Account:</strong> $env:USERDOMAIN\$($env:USERNAME) on $($env:COMPUTERNAME)<br>" +
            "<strong>Module:</strong> $ReporModuleName version $ReportModuleVersion" +
            "</p>"

            $ColumnHeaderBackgroundColor = '#A5D7D2'
            $Head = "<style>" +
            "body {background-color:white; font-family:Calibri; font-size:11pt;}" +
            "p {font-size:10pt; margin-bottom:10;}" +
            "p.footer {font-size:10pt; margin-bottom:10; margin-top:40px;}" +
            "h2 {font-family:Calibri;}" +
            "h3 {font-family:Calibri; margin-top:40px;}" +
            "table {border-width:1px; border-style:solid; border-color:black; border-collapse:collapse;}" +
            "th {border-width:1px; padding:2px; border-style:solid; border-color:black; background-color:$ColumnHeaderBackgroundColor;}" +
            "td {border-width:1px; padding-right:2px; padding-left:2px; padding-top:0px; padding-bottom:0px; border-style:solid; border-color:black; background-color:white;}" +
            "td.HeaderColumn {font-weight: bold; background-color:$ColumnHeaderBackgroundColor;}" +
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

            $Subfolders = Get-PSFConfig -FullName  "$ModuleName.Folder.*"

            foreach ($Subfolder in $Subfolders)
            {
                $SubfolderPath     = $Path + "\{0}" -f (Get-PSFConfigValue -FullName $Subfolder.FullName)
                $SubfolderVariable = $Subfolder.FullName.Split('.')[-1]
                New-Variable -Name ('ScriptPath{0}' -f $SubfolderVariable) -Value $SubfolderPath

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

            #Region Domain Lookup Table
            $DomainSidLookupTable = (Get-ADForest).Domains | Get-ADDomain | ForEach-Object -Begin { $ht = @{} } -Process {
                $ht[$PSItem.DomainSID.Value] = $PSItem
            } -End { return $ht }
            #EndRegin
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

                    foreach ($SearchRootItem in @($SearchRoot))
                    {
                        $ADGroupParams = @{ }

                        # SearchRoot parameter specified
                        if ($SearchRootItem)
                        {
                            Write-Verbose -Message "    [i] [$SearchRootItemNumber] SearchRoot: $SearchRootItem"
                            $ADGroupParams.SearchBase = $SearchRootItem
                        }
                        else
                        {
                            # when no values specified for SearchRoot parameter
                            Write-Verbose -Message "    [i] [$SearchRootItemNumber] Search in whole AD domain"
                        }

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
                                $null = $GroupList.Add($GroupSearchItem)
                            }
                        }

                        $SearchRootItemNumber++
                    }
                }
                'LDAPFilter'
                {
                    Write-Verbose -Message "[*] Using ParameterSet 'LDAPFilter'"

                    $SearchRootItemNumber = 1

                    foreach ($SearchRootItem in @($SearchRoot))
                    {
                        $ADGroupParams = @{ }

                        # SearchRoot parameter specified
                        if ($SearchRootItem)
                        {
                            Write-Verbose -Message "    [i] [$SearchRootItemNumber] SearchRoot: $SearchRootItem"
                            $ADGroupParams.SearchBase = $SearchRootItem
                        }
                        else
                        {
                            # when no values specified for SearchRoot parameter
                            Write-Verbose -Message "    [i] [$SearchRootItemNumber] Search in whole AD domain"
                        }

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

                        # LDAPFilter parameter specified
                        if ($PSBoundParameters['LDAPFilter'])
                        {
                            Write-Verbose -Message "    [i] [$SearchRootItemNumber] LDAPFilter: $LDAPFilter"

                            $ADGroupParams.LDAPFilter = $LDAPFilter
                        }
                        else
                        {
                            # probably obsolote since this ParameterSet is only triggered when a LDAPFilter is provided
                            $ADGroupParams.LDAPFilter = '(objectClass=*)'
                        }

                        $GroupSearch = Get-ADGroup @ADGroupParams -ErrorAction Stop

                        if ($GroupSearch)
                        {
                            foreach ($GroupSearchItem in $GroupSearch)
                            {
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
                $Changes, $ADGroup, $ChangeList, $ChangeHistoryList, $MemberList, $ReferenceObject, $DifferenceObject = $null
                $PreContent, $HtmlGroupSummary, $HtmlChangeList, $HtmlChangeHistoryList, $HtmlMemberList = $null

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
                    if ($ADGroup.SID.AccountDomainSid.value)
                    {
                        $Server    = $DomainSidLookupTable[$ADGroup.SID.AccountDomainSid.value].DNSRoot
                        $DomainName = (Get-ADDomain -Identity $ADGroup.SID.AccountDomainSid.value -Server $Server -ErrorAction Stop).DNSRoot
                    }
                    elseif ($PSBoundParameters['Server'])
                    {
                        $DomainName = (Get-ADDomain -Identity $Server -Server $Server -ErrorAction Stop).DNSRoot
                    }
                    else
                    {
                        $DomainName = (Get-ADDomain -ErrorAction Stop).DNSRoot
                    }

                    Write-Verbose -Message "    [i] Domain: $DomainName"

                    $ADGroupName = $ADGroup.Name

                    #Region Get Group Membership
                    $GroupMemberSplatting = @{}

                    if ($PSBoundParameters['Server'])
                    {
                        $GroupMemberSplatting.Server = $Server
                    }

                    if ($PSBoundParameters['Recursive'])
                    {
                        # Collect direct group membeship to determine value for 'Nested' property
                        $ADGroupMembersDirect = Get-ADGroupMember -Identity $ADGroup @GroupMemberSplatting -ErrorAction Stop -ErrorVariable ErrorProcessGetADGroupMember

                        $GroupMemberSplatting.Recursive = $true
                    }

                    $GroupMemberSplatting.Identity = $ADGroup

                    $ADGroupMembers = Get-ADGroupMember @GroupMemberSplatting -ErrorAction Stop -ErrorVariable ErrorProcessGetADGroupMember

                    Write-Verbose -Message "    [+] Collected AD Group Membership"

                    # For each group member, get AD object with needed properties
                    $Members = $ADGroupMembers | Get-ADObjectDetails -ErrorAction Continue

                    if ($Members)
                    {
                        $ReferenceObject = $Members

                        Write-Verbose -Message ("    [i] AD Group has {0} members" -f @($Members).Count)
                    }
                    else
                    {
                        $ReferenceObject = @()
                        Write-Verbose -Message "    [i] AD Group is empty"
                    }
                    #EndRegion

                    #Region Find Changes

                    # Current Group Membership File
                    # if the file doesn't exist, assume we don't have a record to refer to
                    $CurrentGroupMembershipFilePath = Join-Path -Path $ScriptPathCurrent -ChildPath ("{0}_{1}-Membership.csv" -f $DomainName, $ADGroupName)

                    if (-not (Test-Path -Path $CurrentGroupMembershipFilePath))
                    {
                        Write-Verbose -Message "    [i] CSV file with current group membership does not yet exist"

                        $Members | Select-Object -Property $MembershipCsvProperty |
                            Export-Csv -Path $CurrentGroupMembershipFilePath -NoTypeInformation -Encoding Unicode -ErrorAction Stop
                        Write-Verbose -Message "    [+] Exported current group membership to CSV file '$CurrentGroupMembershipFilePath'"
                    }
                    else
                    {
                        Write-Verbose -Message "    [i] CSV file with current group membership exists"
                    }

                    $MembersFromCsv = Import-Csv -Path $CurrentGroupMembershipFilePath -ErrorAction Stop -ErrorVariable ErrorProcessImportCSV

                    $MembersFromCsvDetails = foreach ($MembersFromCsvItem in $MembersFromCsv)
                    {
                        # for each member in Csv file, get AD object with needed properties, if exists
                        $MembersFromCsvItemDetails = $MembersFromCsvItem | Get-ADObjectDetails -ErrorAction SilentlyContinue -ErrorVariable ADObjectDetailsError

                        if ($ADObjectDetailsError)
                        {
                            $MembersFromCsvItemDetails = [PSCustomObject]@{
                                SamAccountName = $ADObjectDetailsError.TargetObject
                            }

                            Write-Warning ("    [!] Could not get details for AD Object '{0}'" -f $ADObjectDetailsError.TargetObject)
                        }

                        # workaround to combine objects with different types (and properties)
                        # AD Objects and custom PSCustomObject
                        #$MembersFromCsvItemDetails | Select-Object -Property *
                        $MembersFromCsvItemDetails
                    }

                    if ($MembersFromCsvDetails)
                    {
                        $DifferenceObject = $MembersFromCsvDetails
                    }
                    else
                    {
                        $DifferenceObject = @()
                        Write-Verbose -Message "    [i] No current group membership in CSV file"
                    }

                    # compare using the unique identifiert property "SID"
                    $CompareResult = Compare-Object -ReferenceObject $ReferenceObject -DifferenceObject $DifferenceObject -Property SID -PassThru -ErrorAction Stop -ErrorVariable ErrorProcessCompareObject

                    $Changes = $CompareResult | ForEach-Object {
                        $DateTime = Get-Date -Format $ChangesDateTimeFormat
                        $State    = switch ($PSItem.SideIndicator)
                        {
                            '=>' { 'Removed' }
                            '<=' { 'Added' }
                        }

                        $PSItem | Add-Member -Name 'DateTime' -Value $DateTime -MemberType NoteProperty -Force
                        $PSItem | Add-Member -Name 'State' -Value $State -MemberType NoteProperty -Force -PassThru
                    }

                    Write-Verbose -Message "    [+] Compared AD Group Membership with current group membership in CSV file"
                    #EndRegion

                    #if ($Changes -or $AlwaysSendEmail -or $AlwaysSaveReport)
                    if ($Changes -or ($ForceAction -and ($SendEmail -or $SaveReport)))
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
                                $ChangeHistoryItemDetails = $null

                                try
                                {
                                    # For the member in Csv file, get the AD object with needed properties, if exists
                                    #$ChangeHistoryItemDetails = $ChangeHistoryItem | Get-ADObjectDetails -ErrorAction Stop
                                    $ChangeHistoryItemDetails = $ChangeHistoryItem | Get-ADObjectDetails -ErrorAction SilentlyContinue -ErrorVariable ADObjectDetailsError

                                    if ($ADObjectDetailsError)
                                    {
                                        Write-Warning ("    [!] Could not get details for AD Object '{0}'" -f $ADObjectDetailsError.TargetObject)
                                    }
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
                                            if ($ChangeHistoryItemDetails)
                                            {
                                                $PropertyValue = if ($ChangeHistoryItemDetails.Contains($PropertyName)) { $ChangeHistoryItemDetails.$PropertyName.ToString() }
                                                #Write-Verbose "$PropertyName : $PropertyValue"
                                            }
                                            else
                                            {
                                                if ($PropertyName -in ('SamAccountName', 'ObjectClass', 'SID'))
                                                {
                                                    # handling exception where AD object is not found, return the identifier attributes' value from CSV file
                                                    $PropertyValue = $ChangeHistoryItem.$PropertyName
                                                    #Write-Verbose "$PropertyName : $PropertyValue (from CSV)"
                                                }
                                                else
                                                {
                                                    #Write-Verbose "No data in CSV file for property '$PropertyName'"
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
                            Write-Verbose -Message "    [+] Exported current group membership to CSV file '$CurrentGroupMembershipFilePath'"
                        }

                        $MemberList = New-Object System.Collections.Generic.List[Object]

                        foreach ($MemberItem in $Members)
                        {
                            $ListItem = [ordered]@{}

                            foreach ($PropertyName in $MembersTableProperty)
                            {
                                switch ($PropertyName)
                                {
                                    'Nested'
                                    {
                                        # validate if object is direct member in group
                                        $PropertyValue = $null
                                        $PropertyValue = if($MemberItem.SID -notin $ADGroupMembersDirect.SID) { $true }
                                        $ListItem.Add($PropertyName, $PropertyValue)
                                    }
                                    Default
                                    {
                                        $PropertyValue = $null
                                        $PropertyValue = if ($MemberItem.Contains($PropertyName)) { $MemberItem.$PropertyName.ToString() }
                                        $ListItem.Add($PropertyName, $PropertyValue)
                                    }
                                }
                            }

                            $null = $MemberList.Add([PSCustomObject]$ListItem)
                        }

                        Write-Verbose -Message "    [+] Created Member List"
                        #EndRegion

                        #Region Group Summary
                        $GroupSummary = foreach ($PropertyName in $GroupSummaryTableProperty)
                        {
                            $PropertyValue = $null
                            $PropertyValue = if ($ADGroup.Contains($PropertyName)) { $ADGroup.$PropertyName.ToString() }
                            [PSCustomObject]@{
                                Property = $PropertyName
                                Value = $PropertyValue
                            }
                        }
                        #EndRegion

                        #Region HTML Report / Email
                        Write-Verbose -Message "    [*] Preparing the HTML report and Email"

                        # Preparing the body of the email
                        $PreContent = "<h2>Group: $ADGroupName</h2>"

                        if (-not ($PSBoundParameters['ExcludeSummary']))
                        {
                            $HtmlGroupSummaryPreContent = "<h3>Summary</h3>"
                            #$HtmlGroupSummary = $GroupSummary | ConvertTo-Html -PreContent $HtmlGroupSummaryPreContent -Fragment -As List | Out-String
                            $HtmlTable = $GroupSummary | ConvertTo-Html -Fragment
                            [xml]$XML = $HtmlTable

                            foreach ($TableRow in $XML.table.SelectNodes("tr"))
                            {
                                if ($TableRow.th)
                                {
                                    # Remove the header row
                                    $SelectedNodes = $TableRow.SelectNodes("th")
                                    $null = $SelectedNodes | ForEach-Object { $_.ParentNode.RemoveChild($_) }
                                }
                                # Formatting for the table rows: add 'class' to the first cell (<td>)
                                if ($TableRow.td)
                                {
                                    # Tag the TD with a custom 'class'
                                    $TableRow.SelectNodes("td")[0].SetAttribute("class","HeaderColumn")
                                }
                            }

                            $HtmlGroupSummary = ($HtmlGroupSummaryPreContent + $XML.OuterXml) | Out-String
                        }

                        if ($ChangeList -and (-not ($PSBoundParameters['ExcludeChanges'])))
                        {
                            $HtmlChangeListPreContent = "<h3>Membership Changes</h3>"
                            $HtmlChangeListPreContent += "<p><i>The membership of this group changed. See the following Added or Removed members.</i></p>"
                            $HtmlChangeList = $ChangeList | Sort-Object -Property DateTime -Descending | ConvertTo-Html -PreContent $HtmlChangeListPreContent -Fragment | Out-String

                            $MailPriority = 'High'
                        }
                        else
                        {
                            $HtmlChangeListPreContent = "<h3>Membership Changes</h3>"
                            $HtmlChangeListPreContent += "<p><i>No changes since last check</i></p>"
                            $HtmlChangeList = ConvertTo-Html -PreContent $HtmlChangeListPreContent -Fragment | Out-String

                            $MailPriority = 'Normal'
                        }

                        if ($ChangeHistoryList -and (-not ($PSBoundParameters['ExcludeHistory'])))
                        {
                            $HtmlChangeHistoryListPreContent = "<h3>Change History</h3>"
                            $HtmlChangeHistoryListPreContent += "<p><i>List of previous changes on this group observed by the script</i></p>"
                            $HtmlChangeHistoryList = $ChangeHistoryList | Sort-Object -Property DateTime -Descending | ConvertTo-Html -PreContent $HtmlChangeHistoryListPreContent -Fragment | Out-String
                        }

                        if ($MemberList -and $PSBoundParameters['IncludeMembers'])
                        {
                            $HtmlMemberListPreContent = "<h3>Members</h3>"
                            $HtmlMemberListPreContent += "<p><i>List of all members</i></p>"
                            $HtmlMemberList = $MemberList | Sort-Object -Property SamAccountName -Descending | ConvertTo-Html -PreContent $HtmlMemberListPreContent -Fragment | Out-String
                        }

                        $Body = ConvertTo-Html -Head $Head -Body $PreContent, $HtmlGroupSummary, $HtmlChangeList, $HtmlChangeHistoryList, $HtmlMemberList, $PostContent | Out-String
                        $Body = $Body.Replace("Added", "<font color=`"green`"><b>Added</b></font>").Replace("Removed", "<font color=`"red`"><b>Removed</b></font>")

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
                        #elseif ($Changes -or $PSBoundParameters['ForceAction'])
                        elseif ($PSBoundParameters['SendEMail'] -and ($Changes -or $PSBoundParameters['ForceAction']))
                        {
                            # send e-mail for each group if Changes found or if 'ForceAction' is specified

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
                                Encoding   = $EmailEncoding
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

                        #if ($PSBoundParameters['SaveReport'] -or $PSBoundParameters['AlwaysSaveReport'])
                        if ($PSBoundParameters['SaveReport'] -and ($Changes -or $PSBoundParameters['ForceAction']))
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
            #if ($PSBoundParameters['OneReport'] -and ($Changes -or ($PSBoundParameters['SendEmail'] -and $PSBoundParameters['ForceAction'])))
            if ($PSBoundParameters['SendEmail'] -and $PSBoundParameters['OneReport'] -and ($Changes -or $PSBoundParameters['ForceAction']))
            {
                # send OneReport e-mail if 'SendEmail' and 'OneReport' is specified and
                # if Changes found or if 'ForceAction' is specified

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
                    Encoding    = $EmailEncoding
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
