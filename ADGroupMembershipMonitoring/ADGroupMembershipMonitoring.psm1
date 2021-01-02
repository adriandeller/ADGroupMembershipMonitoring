$Script:ModuleRoot = $PSScriptRoot
$Script:ModuleName = Split-Path $PSScriptRoot -Leaf

# Dot source public/private functions
$public  = @(Get-ChildItem -Path (Join-Path -Path $ModuleRoot -ChildPath 'public\*.ps1')  -ErrorAction Stop)
$private = @(Get-ChildItem -Path (Join-Path -Path $ModuleRoot -ChildPath 'private\*.ps1') -ErrorAction Stop)
$config  = @(Get-ChildItem -Path (Join-Path -Path $ModuleRoot -ChildPath 'configuration\*.ps1') -ErrorAction Stop)

foreach ($Import in @($public + $private + $config))
{
    try
    {
        . $Import.FullName
    }
    catch
    {
        throw "Unable to dot source '$($Import.FullName)'"
    }
}

Export-ModuleMember -Function $public.Basename
