function Convert-DateTimeString
{
    [CmdletBinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        [Alias('DateTimeString,')]
        $String,

        [Parameter(Mandatory = $false)]
        [string[]]
        $InputFormat,

        [Parameter(Mandatory = $true)]
        [string]
        $OutputFormat
    )

    begin
    {
    }
    process
    {
        try
        {
            $Result = Get-Date $DateTimeString -Format $OutputFormat -ErrorAction Stop
            Write-Verbose "Successful usging 'Get-Date'"
        }
        catch
        {
            Write-Verbose "Failed format usging 'Get-Date'"
            $Result = $false
        }

        if ($Result)
        {
            return $Result
        }
        else
        {
            foreach ($Format in $InputFormat)
            {
                try
                {
                    $Result = [datetime]::ParseExact($DateTimeString, $Format, $null)
                    Write-Verbose "Successful format: $Format"
                    break
                }
                catch
                {
                    Write-Verbose "Failed format: $Format"
                    $Result = $false
                }

                <#
                [ref]$Result = Get-Date

                [DateTime]::TryParseExact(
                    $DateTimeString,
                    $Format,
                    [System.Globalization.DateTimeFormatInfo]::InvariantInfo,
                    [System.Globalization.DateTimeStyles]::None,
                    $Result
                )

                if ($Result)
                {
                    return (Get-Date -Date $Result)
                }#>
            }

            if ($Result)
            {
                return (Get-Date -Date $Result -Format $OutputFormat)
            }
        }
    }
    end
    {
    }
}
