break

Install-Module -Name platyPS -Scope CurrentUser
Import-Module -Name platyPS

$ModuleName = 'ADGroupMembershipMonitoring'

Import-Module -Name .\$ModuleName -Force
Get-Module -Name $ModuleName
Get-Command -Module $ModuleName

# Create help
New-MarkdownHelp -Module $ModuleName -OutputFolder .\docs -WithModulePage
New-MarkdownHelp -Module $ModuleName -OutputFolder .\docs #-Force
New-ExternalHelp -Path .\docs -OutputPath .\$ModuleName\en-US\

# Update
Update-MarkdownHelpModule -Path .\docs
New-ExternalHelp -Path .\docs -OutputPath .\$ModuleName\en-US\ -Force

# Testing Get-Help for module function
Get-HelpPreview -Path .\$ModuleName\en-US\*.xml
Get-Help Invoke-ADGroupMembershipMonitoring
Get-Help Invoke-ADGroupMembershipMonitoring -Examples
