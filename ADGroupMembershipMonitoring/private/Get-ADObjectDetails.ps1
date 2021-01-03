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
        $DomainSidLookupTable = (Get-ADForest).Domains | Get-ADDomain | ForEach-Object -Begin { $ht = @{} } -Process {
            $ht[$PSItem.DomainSID.Value] = $PSItem
        } -End { return $ht }
    }
    Process
    {
        #Write-Verbose ("ParameterSetName: {0}" -f $PSCmdlet.ParameterSetName)

        if ($PSCmdlet.ParameterSetName -eq 'ADObject')
        {
            # Get Domain (FQDN) of the AD object by the Domain SID
            $DomainSid = $ADObject.SID.AccountDomainSid.Value
            $Server    = $DomainSidLookupTable[$DomainSid].DNSRoot

            $SamAccountName = $ADObject.SamAccountName
            $SID            = $ADObject.SID.Value
            $ObjectClass    = $ADObject.ObjectClass
        }
        else
        {
            # collect required AD object properties (objectClass and SID) for further AD query

            # Get Global Catalog DC
            #$ServerGC = (Get-ADForest).GlobalCatalogs | Get-Random -Count 1
            $ServerGC = (Get-ADDomainController -Discover -Service GlobalCatalog).Hostname
            $Server = '{0}:3268' -f $ServerGC.Value

            try
            {
                if ([string]::IsNullOrEmpty($PSBoundParameters['ObjectClas']))
                {
                    $ADObjectResult = @(Get-ADObject -Filter "SamAccountName -eq '$SamAccountName'" -Server $Server -Property objectSid)

                    Write-Verbose "SID: $SID"

                    if ($ADObjectResult.Count -eq 1)
                    {
                        $ObjectClass = $ADObjectResult.ObjectClass
                        $SID         = $ADObjectResult.objectSid.Value
                    }
                    elseif ($ADObjectResult.Count -gt 1)
                    {
                        $ADObjectResult = ($ADObjectResult | Where-Object { $_.objectSid.Value -eq $SID})
                        $ObjectClass    = $ADObjectResult.ObjectClass
                        $SID            = $ADObjectResult.objectSid.Value

                        if ($ObjectClass)
                        {
                            Write-Verbose "Found AD object by SID '$SID'"
                        }
                        else
                        {
                            Write-Error "Found multiple AD objects with SamAccountName '$SamAccountName'"
                            return
                        }
                    }

                    if ($ObjectClass)
                    {
                        Write-Verbose "Found AD object with ObjectClass '$ObjectClass'"
                    }
                    else
                    {
                        Write-Error "Unable to identify ObjectClass of '$SamAccountName'"
                        return
                    }
                }
            }
            catch
            {
                $_.Exception.Message
            }
        }

        $ADObjectSplatting = @{
            Filter      = "SID -eq '$SID'"
            Server      = $Server
            ErrorAction = 'SilentlyContinue'
        }

        $ADObjectReturn = switch ($ObjectClass)
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
            #throw "AD object '$SamAccountName' not found"

            #
            if ($PSBoundParameters['ErrorAction'] -eq 'Stop')
            {
                throw "AD object '$SamAccountName' not found"
            }
            else
            {
                return [PSCustomObject]@{
                    SamAccountName = $SamAccountName
                    #ObjectClass    = $ObjectClass
                }
            }
            #
        }
    }
    End
    {
    }
}
