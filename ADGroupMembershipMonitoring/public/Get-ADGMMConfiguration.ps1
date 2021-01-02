function Get-ADGMMConfiguration
{
    [CmdletBinding()]

    param
    (
    )

    Begin
    {

    }
    Process
    {
        try
        {
            Get-PSFConfig -Module $Script:ModuleName | Select-Object -Property Module,Name,Value
        }
        catch {
            throw $_.Exception.Message
        }
    }
    End
    {

    }
}
