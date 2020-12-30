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
            $Result = Get-Date $String -Format $OutputFormat -ErrorAction Stop
            Write-Verbose "Successful parsing usging 'Get-Date'"
        }
        catch
        {
            Write-Verbose "Failed parsing using 'Get-Date'"
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
                    $Result = [datetime]::ParseExact($String, $Format, $null)
                    Write-Verbose "Successful parsing using format '$Format'"
                    break
                }
                catch
                {
                    Write-Verbose "Failed parsing using format '$Format'"
                    $Result = $false
                }

                <#
                [ref]$Result = Get-Date

                [DateTime]::TryParseExact(
                    $String,
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
