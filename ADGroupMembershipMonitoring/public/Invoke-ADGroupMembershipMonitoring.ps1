# Requires -Version 5.0

function Invoke-ADGroupMembershipMonitoring
{
    <#
    .DESCRIPTION
        This script is monitoring group(s) in Active Directory and send an email when someone is changing the membership.
        It will also report the Change History made for this/those group(s).
    .SYNOPSIS
        This script is monitoring group(s) in Active Directory and send an email when someone is changing the membership.
    .PARAMETER Group
        Specify the group(s) to query in Active Directory.
        You can also specify the 'DN','GUID','SID' or the 'Name' of your group(s).
        Using 'Domain\Name' will also work.
.PARAMETER Recurse
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
        .\AD-GROUP-Monitor_MemberShip.ps1 -Group "FXGroup" -EmailFrom "From@Company.com" -EmailTo "To@Company.com" -EmailServer "mail.company.com"
        This will run the script against the group FXGROUP and send an email to To@Company.com using the address From@Company.com and the server mail.company.com.
    .EXAMPLE
        .\AD-GROUP-Monitor_MemberShip.ps1 -Group "FXGroup","FXGroup2","FXGroup3" -EmailFrom "From@Company.com" -Emailto "To@Company.com" -EmailServer "mail.company.com"
        This will run the script against the groups FXGROUP,FXGROUP2 and FXGROUP3  and send an email to To@Company.com using the address From@Company.com and the Server mail.company.com.
    .EXAMPLE
        .\AD-GROUP-Monitor_MemberShip.ps1 -Group "FXGroup" -EmailFrom "From@Company.com" -Emailto "To@Company.com" -EmailServer "mail.company.com" -Verbose
        This will run the script against the group FXGROUP and send an email to To@Company.com using the address From@Company.com and the server mail.company.com. Additionally the switch Verbose is activated to show the activities of the script.
    .EXAMPLE
        .\AD-GROUP-Monitor_MemberShip.ps1 -Group "FXGroup" -EmailFrom "From@Company.com" -Emailto "Auditor@Company.com","Auditor2@Company.com" -EmailServer "mail.company.com" -Verbose
        This will run the script against the group FXGROUP and send an email to Auditor@Company.com and Auditor2@Company.com using the address From@Company.com and the server mail.company.com. Additionally the switch Verbose is activated to show the activities of the script.
    .EXAMPLE
        .\AD-GROUP-Monitor_MemberShip.ps1 -SearchRoot 'FX.LAB/TEST/Groups' -Emailfrom Reporting@fx.lab -Emailto "Catfx@fx.lab" -EmailServer 192.168.1.10 -Verbose
        This will run the script against all the groups present in the CanonicalName 'FX.LAB/TEST/Groups' and send an email to catfx@fx.lab using the address Reporting@fx.lab and the server 192.168.1.10. Additionally the switch Verbose is activated to show the activities of the script.
    .EXAMPLE
        .\AD-GROUP-Monitor_MemberShip.ps1 -file .\groupslist.txt -Emailfrom Reporting@fx.lab -Emailto "Catfx@fx.lab" -EmailServer 192.168.1.10 -Verbose
        This will run the script against all the groups present in the file groupslists.txt and send an email to catfx@fx.lab using the address Reporting@fx.lab and the server 192.168.1.10. Additionally the switch Verbose is activated to show the activities of the script.
    .EXAMPLE
        .\AD-GROUP-Monitor_MemberShip.ps1 -server DC01.fx.lab -file .\groupslist.txt -Emailfrom Reporting@fx.lab -Emailto "Catfx@fx.lab" -EmailServer 192.168.1.10 -Verbose
        This will run the script against the Domain Controller "DC01.fx.lab" on all the groups present in the file groupslists.txt and send an email to catfx@fx.lab using the address Reporting@fx.lab and the server 192.168.1.10. Additionally the switch Verbose is activated to show the activities of the script.
    .EXAMPLE
        .\AD-GROUP-Monitor_MemberShip.ps1 -server DC01.fx.lab:389 -file .\groupslist.txt -Emailfrom Reporting@fx.lab -Emailto "Catfx@fx.lab" -EmailServer 192.168.1.10 -Verbose
        This will run the script against the Domain Controller "DC01.fx.lab" (on port 389) on all the groups present in the file groupslists.txt and send an email to catfx@fx.lab using the address Reporting@fx.lab and the server 192.168.1.10. Additionally the switch Verbose is activated to show the activities of the script.
    .EXAMPLE
        .\AD-GROUP-Monitor_MemberShip.ps1 -group "Domain Admins" -Emailfrom Reporting@fx.lab -Emailto "Catfx@fx.lab" -EmailServer 192.168.1.10 -EmailEncoding 'ASCII' -SaveAsHTML
        This will run the script against the group "Domain Admins" and send an email (using the encoding ASCII) to catfx@fx.lab using the address Reporting@fx.lab and the SMTP server 192.168.1.10. It will also save a local HTML report under the HTML Directory.
    .EXAMPLE
        .\AD-GROUP-Monitor_MemberShip.ps1 -Group "FXGroup" -EmailFrom "From@Company.com" -EmailTo "To@Company.com" -EmailServer "mail.company.com" -EmailCredential (Get-Credential)
        This will run the script against the group FXGROUP and send an email to To@Company.com using the address From@Company.com and the server mail.company.com, using the credential specified.
    .INPUTS
        System.String
    .OUTPUTS
        Email Report
    .NOTES
        NAME:    AD-GROUP-Monitor_MemberShip.ps1
        AUTHOR: Adrian Deller
        EMAIL:  adrian.deller@unibas.ch
    #>

    [CmdletBinding(DefaultParameterSetName = "Group")]

    param
    (
        [Parameter(ParameterSetName = "Group", Mandatory = $true, HelpMessage = "You must specify at least one Active Directory group")]
        [ValidateNotNull()]
        [Alias('DN', 'DistinguishedName', 'GUID', 'SID', 'Name')]
        [string[]]
        $Group,

        [Parameter(Mandatory = $false, HelpMessage = "Should the AD group members be searched recursively?")]
        [switch]
        $Recurse,

        [Parameter(ParameterSetName = "OU", Mandatory = $true)]
        [Alias('SearchBase')]
        [string[]]
        $SearchRoot,

        [Parameter(ParameterSetName = "OU")]
        [ValidateSet("Base", "OneLevel", "Subtree")]
        [string]
        $SearchScope,

        [Parameter(ParameterSetName = "OU")]
        [ValidateSet("Global", "Universal", "DomainLocal")]
        [String]
        $GroupScope,

        [Parameter(ParameterSetName = "OU")]
        [ValidateSet("Security", "Distribution")]
        [string]
        $GroupType,

        [Parameter(ParameterSetName = "OU", Mandatory = $true)]
        [string]
        $GroupFilter,

        [Parameter(ParameterSetName = "File", Mandatory = $true)]
        [ValidateScript( { Test-Path -Path $_ })]
        [string[]]
        $File,

        [Parameter()]
        [Alias('DomainController', 'Service')]
        [string]
        $Server,

        [Parameter(Mandatory = $true, HelpMessage = "You must specify the Sender E-Mail Address")]
        [ValidatePattern("[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?")]
        [string]
        $EmailFrom,

        [Parameter(Mandatory = $true, HelpMessage = "You must specify the Destination E-Mail Address")]
        [ValidatePattern("[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?")]
        [string[]]
        $EmailTo,

        [Parameter(Mandatory = $true, HelpMessage = "You must specify the E-Mail Server to use (IPAddress or FQDN)")]
        [string]
        $EmailServer,

        [Parameter(Mandatory = $true, HelpMessage = "You must specify the E-Mail Server Port")]
        [string]
        $EmailPort,

        [Parameter(Mandatory = $false, HelpMessage = "You can provide a prefix for the E-Mail subject")]
        [string]
        $EmailSubjectPrefix,

        [Parameter()]
        [ValidateSet("ASCII", "UTF8", "UTF7", "UTF32", "Unicode", "BigEndianUnicode", "Default")]
        [string]
        $EmailEncoding = "ASCII",

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

        [Parameter(Mandatory = $true, HelpMessage = "You must specify a path for data storage")]
        [Alias('Path', 'OutputPath')]
        [string]
        $Path,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $EmailCredential = [System.Management.Automation.PSCredential]::Empty
    )

    Begin
    {
        try
        {
            # Retrieve the Script name
            $ScriptName = $MyInvocation.MyCommand

            # Create the folders if not present
            if (-not(Test-Path -Path $Path))
            {
                Write-Verbose -Message "[$ScriptName][Begin] Creating the folder for data storage: $ScriptPathCurrent"
                New-Item -Path $ScriptPathCurrent -ItemType Directory | Out-Null
            }

            $ScriptPathCurrent = $Path + "\Current"

            if (-not(Test-Path -Path $ScriptPathCurrent))
            {
                Write-Verbose -Message "[$ScriptName][Begin] Creating the 'Current' Folder: $ScriptPathCurrent"
                New-Item -Path $ScriptPathCurrent -ItemType Directory | Out-Null
            }

            $ScriptPathHistory = $Path + "\History"

            if (-not(Test-Path -Path $ScriptPathHistory))
            {
                Write-Verbose -Message "[$ScriptName][Begin] Creating the 'History' directory: $ScriptPathHistory"
                New-Item -Path $ScriptPathHistory -ItemType Directory | Out-Null
            }

            $ScriptPathHTML = $Path + "\HTML"

            if (-not(Test-Path -Path $ScriptPathHTML))
            {
                Write-Verbose -Message "[$ScriptName][Begin] Creating the 'HTML' directory: $ScriptPathHTML"
                New-Item -Path $ScriptPathHTML -ItemType Directory | Out-Null
            }

            $ScriptPathOneReport = $Path + "\OneReport"

            if (-not(Test-Path -Path $ScriptPathOneReport))
            {
                Write-Verbose -Message "[$ScriptName][Begin] Creating the 'HTML' directory: $ScriptPathOneReport"
                New-Item -Path $ScriptPathOneReport -ItemType Directory | Out-Null
            }

            # Set the Date and Time formats
            $ReporDateFormat = 'yyyy-MM-dd HH:mm:ss'
            $ChangesDateFormat = 'yyyy-MM-dd-HH:mm:ss' #'yyyyMMddHH:mm:ss'
            $FileNameDateFormat = 'yyyyMMdd_HHmmss'

            # Active Directory Module
            Write-Verbose -Message "[$ScriptName][Begin] Active Directory Module"

            # Verify AD Module is loaded
            if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue -ErrorVariable ErrorBeginGetADModule))
            {
                Write-Verbose -Message "[$ScriptName][Begin] Active Directory Module - Loading"
                Import-Module -Name ActiveDirectory -ErrorAction Stop -ErrorVariable ErrorBeginAddADModule
                Write-Verbose -Message "[$ScriptName][Begin] Active Directory Module - Loaded"
            }
            else
            {
                Write-Verbose -Message "[$ScriptName][Begin] Active Directory module seems loaded"
            }

            Write-Verbose -Message "[$ScriptName][Begin] Setting HTML Variables"

            # HTML Report Settings
            $Report = "<p style=`"background-color: white; font-family: consolas; font-size: 9pt;`">" +
            "<strong>Report Time:</strong> $(Get-Date -Format $ReporDateFormat) <br>" +
            "<strong>Account:</strong> $env:USERDOMAIN\$($env:USERNAME) on $($env:COMPUTERNAME)" +
            "</p>"

            $Head = "<style>" +
            "body { background-color: white; font-family: consolas; font-size: 11pt; }" +
            "table { border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; }" +
            "th { border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color:`"#00297a`"; font-color: white; }" +
            "td { border-width: 1px; padding-right: 2px; padding-left: 2px; padding-top: 0px; padding-bottom: 0px; border-style: solid; border-color: black; background-color: white; }" +
            "</style>"

            $Head2 = "<style>" +
            "body { background-color: white; font-family: consolas; font-size: 9pt; }" +
            "table { border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; }" +
            "th { border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: `"#c0c0c0`"; }" +
            "td { border-width: 1px; padding-right: 2px; padding-left: 2px; padding-top: 0px; padding-bottom: 0px; border-style: solid; border-color: black; background-color: white; }" +
            "</style>"
        }
        catch
        {
            Write-Warning -Message "[$ScriptName][Begin] Something went wrong"

            if ($ErrorBeginGetADModule) { Write-Warning -Message "[$ScriptName][Begin] Can't find the Active Directory Module" }
            if ($ErrorBeginAddADModule) { Write-Warning -Message "[$ScriptName][Begin] Can't load the Active Directory Module" }
        }
    }
    Process
    {
        try
        {
            # SearchRoot parameter specified
            if ($PSBoundParameters['SearchRoot'])
            {
                Write-Verbose -Message "[$ScriptName][Process] SearchRoot specified"

                foreach ($item in $SearchRoot)
                {
                    # ADGroup Splatting
                    $ADGroupParams = @{ }

                    $ADGroupParams.SearchBase = $item

                    # Server Specified
                    if ($PSBoundParameters['Server']) { $ADGroupParams.Server = $Server }

                    # SearchScope parameter specified
                    if ($PSBoundParameters['SearchScope'])
                    {
                        Write-Verbose -Message "[$ScriptName][Process] SearchScope specified"

                        $ADGroupParams.SearchScope = $SearchScope
                    }

                    # GroupScope parameter specified
                    if ($PSBoundParameters['GroupScope'])
                    {
                        Write-Verbose -Message "[$ScriptName][Process] GroupScope specified"

                        $ADGroupParams.Filter = "GroupScope -eq `'$GroupScope`'"
                    }

                    # GroupType parameter specified
                    if ($PSBoundParameters['GroupType'])
                    {
                        Write-Verbose -Message "[$ScriptName][Process] GroupType specified"

                        if ($ADGroupParams.Filter)
                        {
                            $ADGroupParams.Filter = "$($ADGroupParams.Filter) -and GroupCategory -eq `'$GroupType`'"
                        }
                        else
                        {
                            $ADGroupParams.Filter = "GroupCategory -eq '$GroupType'"
                        }
                    }

                    if ($PSBoundParameters['GroupFilter'])
                    {
                        Write-Verbose -Message "[$ScriptName][Process] GroupFilter specified"

                        if ($ADGroupParams.Filter)
                        {
                            $ADGroupParams.Filter = "$($ADGroupParams.Filter) -and `'$GroupFilter`'"
                        }
                        else
                        {
                            $ADGroupParams.Filter = $GroupFilter
                        }
                    }

                    if (-not($ADGroupParams.Filter))
                    {
                        $ADGroupParams.Filter = "*"
                    }

                    Write-Verbose -Message "[$ScriptName][Process] AD Module - Querying..."

                    # Add the groups to the variable $Group
                    $GroupSearch = Get-ADGroup @ADGroupParams

                    if ($GroupSearch)
                    {
                        $Group += $GroupSearch.DistinguishedName
                        Write-Verbose -Message "[$ScriptName][Process] OU: $item"
                    }
                }
            }

            # File parameter specified
            if ($PSBoundParameters['File'])
            {
                Write-Verbose -Message "[$ScriptName][Process] File"

                foreach ($item in $File)
                {
                    Write-Verbose -Message "[$ScriptName][Process] Loading File: $item"

                    $FileContent = Get-Content -Path $File

                    if ($FileContent)
                    {
                        $Group += Get-Content -Path $File
                    }
                }
            }

            # Group or SearchRoot or File parameters specified
            #foreach ($item in $Group)
            @($Group).ForEach{

                $item = $PSItem

                try
                {
                    Write-Verbose -Message "[$ScriptName][Process] Group: $item..."

                    # Splatting for the AD Group Request
                    $GroupSplatting = @{ }
                    $GroupSplatting.Identity = $item

                    # Group Information
                    Write-Verbose -Message "[$ScriptName][Process] Active Directory Module"

                    if ($PSBoundParameters['Server'])
                    {
                        $GroupSplatting.Server = $Server
                    }

                    # Look for Group
                    $GroupName = Get-ADGroup @GroupSplatting -Properties * -ErrorAction Continue -ErrorVariable ErrorProcessGetADGroup
                    Write-Verbose -Message "[$ScriptName][Process] Extracting Domain Name from $($GroupName.CanonicalName)"
                    $DomainName = ($GroupName.CanonicalName -split '/')[0]
                    $RealGroupName = $GroupName.Name

                    if ($GroupName)
                    {
                        $GroupMemberSplatting = @{ }
                        $GroupMemberSplatting.Identity = $GroupName

                        # Get GroupName Membership
                        Write-Verbose -Message "[$ScriptName][Process] Group: $item - Querying AD Group Membership"

                        # Add the Server if specified
                        if ($PSBoundParameters['Server'])
                        {
                            $GroupMemberSplatting.Server = $Server
                        }


                        if ($PSBoundParameters['Recurse'])
                        {
                            $MemberObjs = Get-ADGroupMember @GroupMemberSplatting -Recursive -ErrorAction Stop -ErrorVariable ErrorProcessGetADGroupMember
                        }
                        else
                        {
                            $MemberObjs = Get-ADGroupMember @GroupMemberSplatting -ErrorAction Stop -ErrorVariable ErrorProcessGetADGroupMember
                        }

                        [array]$Members = $MemberObjs | Where-Object "$($_.objectClass) -eq 'user' -or $($_.objectClass) -eq 'computer'" | Get-ADObject -Property PasswordExpired | Select-Object -Property *, @{ Name = 'DN'; Expression = { $_.DistinguishedName } }
                        #[Array]$Members = $MemberObjs | Where-Object { $_.objectClass -eq "user" } | Get-ADUser -Properties PasswordExpired | Select-Object -Property *, @{ Name = 'DN'; Expression = { $_.DistinguishedName } }
                        #$Members += $MemberObjs | Where-Object { $_.objectClass -eq "computer" } | Get-ADComputer -Properties PasswordExpired | Select-Object -Property *, @{ Name = 'DN'; Expression = { $_.DistinguishedName } }

                        # no members, add some info in $members to avoid the $null
                        # if the value is $null the compare-object won't work
                        if (-not ($Members))
                        {
                            Write-Verbose -Message "[$ScriptName][Process] Group: $item is empty"

                            $Members = [PSCustomObject]@{
                                Name           = "No User or Group"
                                SamAccountName = "No User or Group"
                            }
                        }

                        # GroupName Membership File
                        # if the file doesn't exist, assume we don't have a record to refer to
                        $StateFile = "$($DomainName)_$($RealGroupName)-membership.csv"

                        if (-not (Test-Path -Path (Join-Path -Path $ScriptPathCurrent -ChildPath $StateFile)))
                        {
                            Write-Verbose -Message "[$ScriptName][Process] $item - The following file did not exist: $StateFile"
                            Write-Verbose -Message "[$ScriptName][Process] $item - Exporting the current membership information into the file: $StateFile"

                            $Members | Export-Csv -Path (Join-Path -Path $ScriptPathCurrent -ChildPath $StateFile) -NoTypeInformation -Encoding Unicode
                        }
                        else
                        {
                            Write-Verbose -Message "[$ScriptName][Process] $item - The following file exists: $StateFile"
                        }

                        # GroupName Membership File is compared with the current GruopName Membership
                        Write-Verbose -Message "[$ScriptName][Process] $item - Comparing Current and Before"

                        $ImportCSV = Import-Csv -Path (Join-Path -Path $ScriptPathCurrent -ChildPath $StateFile) -ErrorAction Stop -ErrorVariable ErrorProcessImportCSV

                        $Changes = Compare-Object -DifferenceObject $ImportCSV -ReferenceObject $Members -ErrorAction Stop -ErrorVariable ErrorProcessCompareObject -Property Name, SamAccountName, DN |
                            Select-Object -Property @{ Name = 'DateTime'; Expression = { Get-Date -Format } }, @{
                                Name = 'State'; expression = {
                                    if ($_.SideIndicator -eq "=>")
                                    {
                                        "Removed"
                                    }
                                    else
                                    {
                                        "Added"
                                    }
                                }
                            }, DisplayName, Name, SamAccountName, DN | Where-Object { $_.Name -notlike "*no user or group*" }

                        Write-Verbose -Message "[$ScriptName][Process] $item - Compare Block Done!"

                        <# Troubleshooting
                        Write-Verbose -Message "[$ScriptName]IMPORTCSV var"
                        $ImportCSV | fl -Property Name, SamAccountName, DN

                        Write-Verbose -Message "[$ScriptName]MEMBER"
                        $Members | fl -Property Name, SamAccountName, DN

                        Write-Verbose -Message "[$ScriptName]CHANGE"
                        $Changes
                        #>

                        # Changes Found
                        if ($Changes -or $AlwaysReport)
                        {
                            Write-Verbose -Message "[$ScriptName][Process] $item - Some changes found"
                            $Changes | Select-Object -Property DateTime, State, Name, SamAccountName, DN

                            # Change History
                            # Get the Past Changes History
                            Write-Verbose -Message "[$ScriptName][Process] $item - Get the change history for this group"
                            $ChangeHistoryFiles = Get-ChildItem -Path $ScriptPathHistory\$($DomainName)_$($RealGroupName)-ChangeHistory.csv -ErrorAction SilentlyContinue
                            Write-Verbose -Message "[$ScriptName][Process] $item - Change history files: $(($ChangeHistoryFiles|Measure-Object).Count)"

                            # Process each history changes
                            if ($ChangeHistoryFiles)
                            {
                                $InfoChangeHistory = @()

                                foreach ($file in $ChangeHistoryFiles.FullName)
                                {
                                    Write-Verbose -Message "[$ScriptName][Process] $item - Change history files - Loading $file"

                                    # Import the file and show the $file creation time and its content
                                    $ImportedFile = Import-Csv -Path $file -ErrorAction Stop -ErrorVariable ErrorProcessImportCSVChangeHistory

                                    foreach ($obj in $ImportedFile)
                                    {
                                        $Output = "" | Select-Object -Property DateTime, State, DisplayName, Name, SamAccountName, DN
                                        $Output.DateTime = $obj.DateTime
                                        $Output.State = $obj.State
                                        $Output.DisplayName = $obj.DisplayName
                                        $Output.Name = $obj.Name
                                        $Output.SamAccountName = $obj.SamAccountName
                                        $Output.DN = $obj.DN
                                        $InfoChangeHistory = $InfoChangeHistory + $Output
                                    } # foreach ($obj in $ImportedFile)
                                } # foreach ($file in $ChangeHistoryFiles.FullName)

                                Write-Verbose -Message "[$ScriptName][Process] $item - Change History Process Completed"
                            } # if ($ChangeHistoryFiles)

                            if ($IncludeMembers)
                            {
                                $InfoMembers = @()
                                Write-Verbose -Message "[$ScriptName][Process] $item - Full Member List - Loading"

                                foreach ($obj in $Members)
                                {
                                    $Output = "" | Select-Object -Property Name, SamAccountName, DN, Enabled, PasswordExpired
                                    $Output.Name = $obj.Name
                                    $Output.SamAccountName = $obj.SamAccountName
                                    $Output.DN = $obj.DistinguishedName

                                    if ($ExtendedProperty)
                                    {
                                        $Output.Enabled = $obj.Enabled
                                        $Output.PasswordExpired = $obj.PasswordExpired
                                    }

                                    $InfoMembers = $InfoMembers + $Output
                                } #foreach ($obj in $Members)

                                Write-Verbose -Message "[$ScriptName][Process] $item - Full Member List Process Completed"
                            } # if ($IncludeMembers)

                            if ($Changes)
                            {
                                Write-Verbose -Message "[$ScriptName][Process] $item - Save Changes to the group's ChangeHistory file"

                                if (-not (Test-Path -Path (Join-Path -Path $ScriptPathHistory -ChildPath "$($DomainName)_$($RealGroupName)-ChangeHistory.csv")))
                                {
                                    $Changes | Export-Csv -Path (Join-Path -Path $ScriptPathHistory -ChildPath "$($DomainName)_$($RealGroupName)-ChangeHistory.csv") -NoTypeInformation -Encoding Unicode
                                }
                                else
                                {
                                    $Changes | Export-Csv -Path (Join-Path -Path $ScriptPathHistory -ChildPath "$($DomainName)_$($RealGroupName)-ChangeHistory.csv") -NoTypeInformation -Append -Encoding Unicode
                                }
                            } # if ($Changes)

                            # Email
                            Write-Verbose -Message "[$ScriptName][Process] $item - Preparing the notification email..."

                            # Preparing the body of the email
                            $Body = "<h2>Group: $($GroupName.SamAccountName)</h2>"
                            $Body += "<p style=`"background-color: white; font-family: consolas; font-size: 8pt;`">"
                            $Body += "<u>Group Description:</u> $($GroupName.Description)<br>"
                            $Body += "<u>Group DistinguishedName:</u> $($GroupName.DistinguishedName)<br>"
                            $Body += "<u>Group CanonicalName:</u> $($GroupName.CanonicalName)<br>"
                            $Body += "<u>Group SID:</u> $($GroupName.Sid.Value)<br>"
                            $Body += "<u>Group Scope/Type:</u> $($GroupName.GroupScope) / $($GroupName.GroupType)<br>"
                            $Body += "</p>"

                            if ($Changes)
                            {
                                $Body += "<h3>Membership Change</h3>"
                                $Body += "<i>The membership of this group changed. See the following Added or Removed members.</i>"

                                # Removing the old DisplayName Property
                                #$Changes = $Changes | Select-Object -Property DateTime, State, Name, SamAccountName, DN
                                $Changes = $Changes | Select-Object -Property @{Name = 'DateTime'; Expression = { Get-Date $_.DateTime -Format $ChangesDateFormat } }, State, Name, SamAccountName, DN

                                $Body += $changes | ConvertTo-Html -head $Head | Out-String
                                $Body += "<br><br><br>"
                            }
                            else
                            {
                                $Body += "<h3>Membership Change</h3>"
                                $Body += "<i>No changes.</i>"
                            }

                            if ($InfoChangeHistory)
                            {
                                # Removing the old DisplayName Property
                                #$InfoChangeHistory = $InfoChangeHistory | Select-Object -Property DateTime, State, Name, SamAccountName, DN
                                $InfoChangeHistory = $InfoChangeHistory | Select-Object -Property @{Name = 'DateTime'; Expression = { Get-Date $_.DateTime -Format $ChangesDateFormat } }, State, Name, SamAccountName, DN

                                $Body += "<h3>Change History</h3>"
                                $Body += "<i>List of previous changes on this group observed by the script</i>"
                                $Body += $InfoChangeHistory | Sort-Object -Property DateTime -Descending | ConvertTo-Html -Fragment -PreContent $Head2 | Out-String
                            }

                            if ($InfoMembers)
                            {
                                $Body += "<h3>Members</h3>"
                                $Body += "<i>List of all members</i>"
                                $Body += $InfoMembers | Sort-Object -Property SamAccountName -Descending | ConvertTo-Html -Fragment -PreContent $Head2 | Out-String
                            }

                            $Body = $Body -replace "Added", "<font color=`"green`"><b>Added</b></font>"
                            $Body = $Body -replace "Removed", "<font color=`"red`"><b>Removed</b></font>"
                            $Body += $Report

                            if ($EmailSubjectPrefix)
                            {
                                $EmailSubject = "$EmailSubjectPrefix $($GroupName.SamAccountName) Membership Change"
                            }
                            else
                            {
                                $EmailSubject = "[AD Group Membership Monitoring] $($GroupName.SamAccountName) Membership Change"
                            }

                            # only send e-mail for each group if 'OneReport' is not specified
                            if (-not ($OneReport))
                            {
                                $mailParam = @{
                                    To         = $EmailTo
                                    From       = $EmailFrom
                                    Subject    = $EmailSubject
                                    Body       = $Body
                                    SmtpServer = $EmailServer
                                    Port       = $EmailPort
                                    Credential = $EmailCredential
                                }

                                Send-MailMessage @mailParam -UseSsl -BodyAsHtml

                                Write-Verbose -Message "[$ScriptName][Process] $item - Email Sent."
                            }

                            # GroupName Membership export to CSV
                            Write-Verbose -Message "[$ScriptName][Process] $item - Exporting the current membership to $StateFile"
                            $Members | Export-Csv -Path (Join-Path -Path $ScriptPathCurrent -ChildPath $StateFile) -NoTypeInformation -Encoding Unicode

                            # Define HTML File Name
                            $HTMLFileName = "$($DomainName)_$($RealGroupName)-$(Get-Date -Format $FileNameDateFormat).html"

                            if ($PSBoundParameters['SaveAsHTML'])
                            {
                                # Save HTML File
                                $Body | Out-File -FilePath (Join-Path -Path $ScriptPathHTML -ChildPath $HTMLFileName)
                            }

                            if ($OneReport)
                            {
                                # Save HTML File
                                $Body | Out-File -FilePath (Join-Path -Path $ScriptPathOneReport -ChildPath $HTMLFileName)
                            }
                        }
                        else
                        {
                            Write-Verbose -Message "[$ScriptName][Process] $item - No Change"
                        }
                    }
                    else
                    {
                        Write-Verbose -Message "[$ScriptName][Process] $item - Group can't be found"
                    }
                } # try
                catch
                {
                    Write-Warning -Message "[$ScriptName][Process] Something went wrong"

                    # Active Directory Module Errors
                    if ($ErrorProcessGetADGroup) { Write-Warning -Message "[$ScriptName][Process] AD Module - Error querying the gruop $item in Active Directory"; }
                    if ($ErrorProcessGetADGroupMember) { Write-Warning -Message "[$ScriptName][Process] AD Module - Error querying the group $item members in Active Directory"; }

                    # Import CSV Errors
                    if ($ErrorProcessImportCSV) { Write-Warning -Message "[$ScriptName][Process] Error importing $StateFile"; }
                    if ($ErrorProcessCompareObject) { Write-Warning -Message "[$ScriptName][Process] Error when comparing"; }
                    if ($ErrorProcessImportCSVChangeHistory) { Write-Warning -Message "[$ScriptName][Process] Error importing $file"; }

                    Write-Warning -Message $_.Exception.Message
                }
            }

            if ($OneReport)
            {
                [string[]]$Attachments = @()

                foreach ($a in (Get-ChildItem $ScriptPathOneReport))
                {
                    $Attachments.Add($a.fullname)
                }

                if ($EmailSubjectPrefix)
                {
                    $EmailSubject = "$EmailSubjectPrefix Membership Report"
                }
                else
                {
                    $EmailSubject = "[AD Group Membership Monitoring] Membership Report"
                }

                $mailParam = @{
                    To          = $EmailTo
                    From        = $EmailFrom
                    Subject     = $EmailSubject
                    Body        = "<h2>See Report in Attachment</h2>"
                    SmtpServer  = $EmailServer
                    Port        = $EmailPort
                    Credential  = $EmailCredential
                    Attachments = $Attachments
                }

                Send-MailMessage @mailParam -UseSsl -BodyAsHtml

                foreach ($a in $Attachments)
                {
                    $a.Dispose()
                }

                Get-ChildItem $ScriptPathOneReport | Remove-Item -Force -Confirm:$false
                Write-Verbose -Message "[$ScriptName][Process] OneReport - Email Sent."
            }
        }
        catch
        {
            Write-Warning -Message "[$ScriptName][Process] Something went wrong"
            throw $_
        }
    }
    End
    {
        Write-Verbose -Message "[$ScriptName][End] Script Completed"
    }
}
