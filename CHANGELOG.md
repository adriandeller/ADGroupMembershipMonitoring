## ADGroupMembershipMonitoring Release History

## 0.1.0 - (2021-xx-xx)

**Added**

- new PowerShell Script Module, main function code is based on script Monitor-ADGroupMemberShip.ps1 version 2.0.7 (2019.08.23)
- add parameter 'Recursive' for indirect membership through group nesting
- add parameter 'EmailSubjectPrefix'
- add parameter 'Path'
- add parameter 'GroupFilter'
- add parameter 'Recursive'
- add Computer, Group and User objects as possible group members for monitoring

**Changed**

- rename parameter 'HTMLLog' to 'SaveAsHTML'

**Removed**

- remove all "Quest Active Directory PowerShell Snapin" related code
