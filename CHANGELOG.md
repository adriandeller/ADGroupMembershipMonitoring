## ADGroupMembershipMonitoring Release History

## 0.1.0 - (2021-xx-xx)

**Added**

- new PowerShell Script Module
- the main function's code is based on script Monitor-ADGroupMemberShip.ps1 version 2.0.7 (2019.08.23)
- add support for Computer, Group and User objects as group members for monitoring
- add parameter 'Path'
- add parameter 'Recursive' for indirect membership through group nesting
- add parameter 'GroupFilter'
- add parameter 'LDAPFilter'
- add parameter 'EmailSubjectPrefix'
- add parameter 'EmailCredential'
- add parameter 'IncludeMembers'
- add parameter 'IncludeManagers
- add parameter 'ExcludeSummary'
- add parameter 'ExcludeChanges'
- add parameter 'ExcludeHistory'
- add parameter 'SendEmail'
- add parameter 'ForceAction'

**Changed**

- rename parameter 'HTMLLog' to 'SaveReport'
- change parameter 'ExtendedProperty' from switch to accept an array of strings
- improve HTML code for tables and style
- change AD attribute used as the unique identifier for AD objects to 'SID'

**Removed**

- remove all "Quest Active Directory PowerShell Snapin" related code
