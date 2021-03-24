# Set defaults
Set-PSFConfig -Module $ModuleName -Name 'Folder.Current' -Value 'Current' -Initialize -Validation string -Description ''
Set-PSFConfig -Module $ModuleName -Name 'Folder.History' -Value 'History' -Initialize -Validation string -Description ''
Set-PSFConfig -Module $ModuleName -Name 'Folder.HTML' -Value 'HTML' -Initialize -Validation string -Description ''
Set-PSFConfig -Module $ModuleName -Name 'Folder.OneReport' -Value 'OneReport' -Initialize -Validation string -Description ''
Set-PSFConfig -Module $ModuleName -Name 'Email.EmailSubjectPrefix' -Value '[AD Group Membership Monitoring]' -Initialize -Validation string -Description ''
Set-PSFConfig -Module $ModuleName -Name 'CSV.ChangeHistoryProperty' -Value 'DateTime', 'State', 'Name', 'SamAccountName', 'SID', 'ObjectClass', 'ObjectGUID' -Initialize -Validation stringarray -Description ''
Set-PSFConfig -Module $ModuleName -Name 'CSV.MembershipProperty' -Value 'Name', 'SamAccountName', 'SID', 'ObjectClass', 'ObjectGUID' -Initialize -Validation stringarray -Description ''
Set-PSFConfig -Module $ModuleName -Name 'HTML.TableHeaderBackgroundColor' -Value '#A5D7D2' -Initialize -Validation string -Description ''

<#
# Gather settings to export
$configToExport = @()
$configToExport += Set-PSFConfig -Module $Module -Name 'Folder.Current' -Value 'Current' -SimpleExport -PassThru
$configToExport += Set-PSFConfig -Module $Module -Name 'Folder.History' -Value 'History' -SimpleExport -PassThru
$configToExport += Set-PSFConfig -Module $Module -Name 'Folder.HTML' -Value 'HTML' -SimpleExport -PassThru
$configToExport += Set-PSFConfig -Module $Module -Name 'Folder.OneReport' -Value 'OneReport' -SimpleExport -PassThru
$configToExport += Set-PSFConfig -Module $Module -Name 'Email.EmailSubjectPrefix' -Value '[AD Group Membership Monitoring]' -SimpleExport -PassThru
$configToExport += Set-PSFConfig -Module $Module -Name 'CSV.ChangeHistoryProperty' -Value 'DateTime', 'State', 'Name', 'SamAccountName', 'ObjectClass' -SimpleExport -PassThru
$configToExport += Set-PSFConfig -Module $Module -Name 'CSV.MembershipProperty' -Value 'Name', 'SamAccountName', 'ObjectClass' -SimpleExport -PassThru


# Write the configuration file
$configToExport | Export-PSFConfig -OutPath (Join-Path -Path $ModuleRoot -ChildPath 'configuration\configuration.json')
#>
