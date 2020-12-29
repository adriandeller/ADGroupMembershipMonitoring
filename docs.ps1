break

Install-Module -Name platyPS -Scope CurrentUser -Force
Import-Module platyPS -Force

$Module = Split-Path -Path $PWD -Leaf

# you should have module imported in the session
Import-Module -Name .\$Module -Force
Get-Command -Module $Module

# Create help
New-MarkdownHelp -Module $Module -OutputFolder .\docs -WithModulePage
New-MarkdownHelp -Module $Module -OutputFolder .\docs #-Force
New-ExternalHelp -Path .\docs -OutputPath .\$Module\en-US\

# Update
Update-MarkdownHelpModule -Path .\docs
New-ExternalHelp -Path .\docs -OutputPath .\$Module\en-US\ -Force

# Testing Get-Help for module function
Get-HelpPreview -Path .\$Module\en-US\*.xml
Get-Help Export-UniADLocalGroupToConfluence
Get-Help Export-UniADLocalGroupToConfluence -Examples
