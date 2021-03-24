function Get-ADObjectDetails
{
    [CmdletBinding(DefaultParameterSetName = 'ADObject')]

    param
    (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Property')]
        [string]
        $SamAccountName,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Property')]
        [AllowNull()]
        [string]
        $SID,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Property')]
        [AllowNull()]
        [string]
        $ObjectClass,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'ADObject')]
        [Microsoft.ActiveDirectory.Management.ADObject]
        $ADObject
    )

    Begin
    {
        <#
        # moved to Begin block in main function
        $DomainSidLookupTable = (Get-ADForest).Domains | Get-ADDomain | ForEach-Object -Begin { $ht = @{} } -Process {
            $ht[$PSItem.DomainSID.Value] = $PSItem
        } -End { return $ht }
        #>
    }
    Process
    {
        Write-Verbose ("ParameterSetName: {0}" -f $PSCmdlet.ParameterSetName)

        $ADObjectToProcess = $null

        if ($PSCmdlet.ParameterSetName -eq 'ADObject')
        {
            # Get Domain (FQDN) of the AD object by the Domain SID
            if ($ADObject.SID)
            {
                $DomainSid = $ADObject.SID.AccountDomainSid.Value
            }
            else
            {
                # ADObject from Get-ADObject has no 'SID' attribute but 'objectSID'
                $ADObject | Add-Member -MemberType AliasProperty -Name 'SID' -Value 'objectSID' -Force
                $DomainSid = $ADObject.objectSID.AccountDomainSid.Value
            }

            Write-Verbose ("DomainSid: {0}" -f $DomainSid)
            $Server    = $DomainSidLookupTable[$DomainSid].DNSRoot

            #$SamAccountName = $ADObject.SamAccountName
            #$SID            = $ADObject.SID.Value
            #$ObjectClass    = $ADObject.ObjectClass

            $ADObjectToProcess = $ADObject
        }
        else
        {
            # collect required AD object properties (objectClass and SID) for further AD query

            Write-Verbose "    Param 'SamAccountName': $SamAccountName"
            Write-Verbose "    Param 'SID': $SID"
            Write-Verbose "    Param 'ObjectClass': $ObjectClass"

            try
            {
                # Get Global Catalog DC
                #$ServerGC = (Get-ADForest).GlobalCatalogs | Get-Random -Count 1
                $ServerGC = (Get-ADDomainController -Discover -Service GlobalCatalog -ErrorAction Stop).Hostname
                $Server = '{0}:3268' -f $ServerGC.Value

                #if ([string]::IsNullOrEmpty($PSBoundParameters['ObjectClass']))
                #{
                    $ADObjectResult = @(Get-ADObject -Filter "SamAccountName -eq '$SamAccountName'" -Server $Server -Property objectSid)

                    if ($ADObjectResult.Count -eq 1)
                    {
                        $ADObjectToProcess = $ADObjectResult

                        if ($ADObjectToProcess.ObjectClass)
                        {
                            Write-Verbose ("    Found AD object with ObjectClass '{0}'" -f $ADObjectToProcess.ObjectClass)
                        }
                    }
                    elseif ($ADObjectResult.Count -gt 1)
                    {
                        $ADObjectResult = ($ADObjectResult | Where-Object { $_.objectSid.Value -eq $SID})
                        $ADObjectToProcess = $ADObjectResult

                        if ($ADObjectToProcess.ObjectClass)
                        {
                            Write-Verbose ("    Found AD object by SID '{0}'" -f $ADObjectToProcess.objectSid.Value)
                        }
                        else
                        {
                            Write-Warning "Found multiple AD objects with SamAccountName '$SamAccountName'"
                        }
                    }

                    if (-not ($ADObjectToProcess.ObjectClass))
                    {
                        $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                            "Could not determine ObjectClass of '$SamAccountName'", # exception
                            $null, # errorId
                            [System.Management.Automation.ErrorCategory]::NotSpecified, # errorCategory
                            $SamAccountName) # targetObject
                          )
                          return
                    }
                #}
            }
            catch
            {
                $_.Exception.Message
            }
        }

        if ($PSCmdlet.ParameterSetName -eq 'ADObject')
        {
            $Filter = "SID -eq '{0}'" -f $ADObjectToProcess.SID.Value
        }
        else
        {
            $Filter = "SID -eq '{0}'" -f $ADObjectToProcess.objectSid.Value
        }

        $ADObjectSplatting = @{
            Filter      = $Filter
            Server      = $Server
            ErrorAction = 'SilentlyContinue'
        }

        $ADObjectReturn = switch ($ADObjectToProcess.ObjectClass)
        {
            'computer' { Get-ADComputer @ADObjectSplatting }
            'group' { Get-ADGroup @ADObjectSplatting }
            'user' { Get-ADUser @ADObjectSplatting }
            #Default { throw "Provided invalid value for 'ObjectClass'"}
        }

        if ($ADObjectReturn)
        {
            return $ADObjectReturn
        }
        else
        {
            if ($PSBoundParameters['ErrorAction'] -eq 'Stop')
            {
                throw "Could not find AD object '$SamAccountName'"
            }
            else
            {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    "Could not find AD object '$SamAccountName'",
                    $null, # error ID
                    [System.Management.Automation.ErrorCategory]::InvalidData, # error category
                    $SamAccountName) # offending object
                  )
            }
        }
    }
    End
    {
    }
}
