---
external help file: ADGroupMembershipMonitoring-help.xml
Module Name: ADGroupMembershipMonitoring
online version:
schema: 2.0.0
---

# Invoke-ADGroupMembershipMonitoring

## SYNOPSIS
This function is monitoring group(s) in Active Directory and sends an email when someone is changing the membership.

## SYNTAX

### Group (Default)
```
Invoke-ADGroupMembershipMonitoring -Group <String[]> [-Recursive] [-Server <String>] -EmailFrom <String>
 -EmailTo <String[]> -EmailServer <String> [-EmailPort <String>] [-EmailSubjectPrefix <String>]
 [-EmailEncoding <String>] [-SaveAsHTML] [-IncludeMembers] [-AlwaysReport] [-OneReport] [-ExtendedProperty]
 -Path <String> [-EmailCredential <PSCredential>] [<CommonParameters>]
```

### OU
```
Invoke-ADGroupMembershipMonitoring [-Recursive] -SearchRoot <String[]> [-SearchScope <String>]
 [-GroupScope <String>] [-GroupType <String>] [-GroupFilter <String>] [-Server <String>] -EmailFrom <String>
 -EmailTo <String[]> -EmailServer <String> [-EmailPort <String>] [-EmailSubjectPrefix <String>]
 [-EmailEncoding <String>] [-SaveAsHTML] [-IncludeMembers] [-AlwaysReport] [-OneReport] [-ExtendedProperty]
 -Path <String> [-EmailCredential <PSCredential>] [<CommonParameters>]
```

### File
```
Invoke-ADGroupMembershipMonitoring [-Recursive] -File <String[]> [-Server <String>] -EmailFrom <String>
 -EmailTo <String[]> -EmailServer <String> [-EmailPort <String>] [-EmailSubjectPrefix <String>]
 [-EmailEncoding <String>] [-SaveAsHTML] [-IncludeMembers] [-AlwaysReport] [-OneReport] [-ExtendedProperty]
 -Path <String> [-EmailCredential <PSCredential>] [<CommonParameters>]
```

## DESCRIPTION
This function is monitoring group(s) in Active Directory and sends an email when someone is changing the membership.
It will also report the Change History made for this/those group(s).

## EXAMPLES

### Example 1
```
PS> Invoke-ADGroupMembershipMonitoring -Group 'Domain Admins' -EmailFrom 'From@Company.com' -EmailTo 'To@Company.com' -EmailServer 'mail.company.com'
```

This will query the group 'Domain Admins' and send an email to 'To@Company.com' using the address 'From@Company.com' and the server 'mail.company.com'.

## PARAMETERS

### -AlwaysReport
{{ Fill AlwaysReport Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmailCredential
{{ Fill EmailCredential Description }}

```yaml
Type: PSCredential
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmailEncoding
You can specify the type of encoding

```yaml
Type: String
Parameter Sets: (All)
Aliases:
Accepted values: ASCII, UTF8, UTF7, UTF32, Unicode, BigEndianUnicode, Default

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmailFrom
You must specify the sender E-Mail Address

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmailPort
You can specify an alternate port on the (SMTP) Mail Server

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmailServer
You must specify the Mail Server to use (FQDN or ip address)

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmailSubjectPrefix
You can provide a prefix for the E-Mail subject

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmailTo
You must specify the destination E-Mail Address

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExtendedProperty
{{ Fill ExtendedProperty Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -File
You must specify at least one file

```yaml
Type: String[]
Parameter Sets: File
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Group
You must specify at least one Active Directory group

```yaml
Type: String[]
Parameter Sets: Group
Aliases: DN, DistinguishedName, GUID, SID, Name

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GroupFilter
{{ Fill GroupFilter Description }}

```yaml
Type: String
Parameter Sets: OU
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GroupScope
{{ Fill GroupScope Description }}

```yaml
Type: String
Parameter Sets: OU
Aliases:
Accepted values: Global, Universal, DomainLocal

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GroupType
{{ Fill GroupType Description }}

```yaml
Type: String
Parameter Sets: OU
Aliases:
Accepted values: Security, Distribution

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeMembers
{{ Fill IncludeMembers Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -OneReport
{{ Fill OneReport Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
You must specify a path for data storage

```yaml
Type: String
Parameter Sets: (All)
Aliases: OutputPath, FolderPath

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Recursive
Should the AD group members be searched recursively?

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -SaveAsHTML
{{ Fill SaveAsHTML Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -SearchRoot
You must specify at least one Active Directory OU

```yaml
Type: String[]
Parameter Sets: OU
Aliases: SearchBase

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SearchScope
{{ Fill SearchScope Description }}

```yaml
Type: String
Parameter Sets: OU
Aliases:
Accepted values: Base, OneLevel, Subtree

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Server
{{ Fill Server Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases: DomainController, Service

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### None
## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS
