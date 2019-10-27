break

Install-Module -Name platyPS -Scope CurrentUser
Import-Module -Name platyPS

$ModuleName = 'ADGroupMembershipMonitoring'

Import-Module -Name .\$ModuleName -Force

Get-Module -Name $ModuleName

Get-Command -Module $ModuleName

New-MarkdownHelp -Module $ModuleName -OutputFolder .\docs -NoMetadata -Force -WithModulePage
New-ExternalHelp -Path .\docs -OutputPath .\$ModuleName\en-US\

# Testing Get-Help for module function
Get-Help Invoke-ADGroupMembershipMonitoring
