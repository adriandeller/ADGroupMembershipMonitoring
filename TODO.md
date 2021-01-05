# To Do

- [ ] add parameter 'LimitHistory' to limit Change History entries in the Report (number entries or timespan ?)
- [ ] add parameter 'LDAPFilter' to gather AD Groups using an LDAP Filter
- [x] add parameter 'AlwaysExport' to create a HTML Report each time
- [x] add parameter 'ExcludeSummary' to exclude the Summary at the top of the Report
- [x] add parameter 'ExcludeHistory' to exclude the Change History in the Report
- [x] add parameter 'ExcludeChanges' to exclude the Changes in the Report
- [x] add parameter 'IncludeMembers' to include the Members in the Report
- [ ] add parameter 'EmailToSelf' to send email to the group's mail-address
- [ ] add parameter 'EmailToManger' send email to the group's manager(s) mail-address (related attributes: managedBy, msExchCoManagedByLink)
- [x] add parameter 'Recursive' for indirect membership through group nesting
- [x] add parameter 'EmailSubjectPrefix'
- [x] add parameter 'Path' to specify a path to store the CSV and HTML files
- [x] add parameter 'GroupFilter'
- [x] add support for Computer and Group object as member of AD Groups
- [ ] add Current Member Count, Added Member count, Removed Member Count
- [ ] add column to indicate if it's a direct/indirect member of the group, when using the 'Recursive' parameter
