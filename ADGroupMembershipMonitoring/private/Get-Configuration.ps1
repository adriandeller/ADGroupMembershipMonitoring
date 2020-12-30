function Get-Configuration
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
            $ModuleRoot = $MyInvocation.MyCommand.Module.ModuleBase
            #Import-PowerShellDataFile -Path "$PSScriptRoot\..\config\configuration.psd1" -ErrorAction Stop
            Import-PowerShellDataFile -Path "$ModuleRoot\config\configuration.psd1" -ErrorAction Stop
        }
        catch {
            throw $_.Exception.Message
        }
    }
    End
    {

    }
}
