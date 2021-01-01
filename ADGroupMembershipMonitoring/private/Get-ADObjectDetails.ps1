function Get-ADObjectDetails
{
    [CmdletBinding()]

    param
    (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $SamAccountName,

        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [string]
        $ObjectClass
    )

    Begin
    {
    }
    Process
    {
        if ([string]::IsNullOrEmpty($ObjectClass))
        {
            try
            {
                $ObjectClass = (Get-ADObject -Filter "SamAccountName -eq '$SamAccountName'").ObjectClass

                if ($ObjectClass)
                {
                    Write-Verbose "Found AD object with 'ObjectClass' = '$ObjectClass'"
                }
                else
                {
                    throw "Unable to identify 'ObjectClass' of '$SamAccountName'"
                }
            }
            catch
            {
                $_.Exception.Message
            }
        }

        $ADObjectSplatting = @{
            Filter          = "SamAccountName -eq '$SamAccountName'"
            ErrorAction = 'SilentlyContinue'
        }

        $ADObject = switch ($ObjectClass)
        {
            'computer' { Get-ADComputer @ADObjectSplatting }
            'group' { Get-ADGroup @ADObjectSplatting }
            'user' { Get-ADUser @ADObjectSplatting }
            Default { throw "Provided invalid value for 'ObjectClass'"}
        }

        if ($ADObject)
        {
            return $ADObject
        }
        else
        {
            throw "AD object '$SamAccountName' not found"

            <#
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
            #>
        }
    }
    End
    {
    }
}
