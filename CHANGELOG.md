# Change Log

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [1.0.0] Unreleased
### Added
* convert script to PowerShell Script Module
* function code is baded on script Monitor-ADGroupMemberShip.ps1 version 2.0.7 (2019.08.23)
* add parameter 'Recursive' for indirect membership through group nesting
* add parameter 'MailSubjectPrefix'
* add parameter 'Path ('OutputPath')
* add parameter 'GroupFilter'
* add User and Computer as possible group members for monitoring

### Changed
* rename parameter 'HTMLLog' to 'SaveAsHTML'

### Removed
* remove all "Quest Active Directory PowerShell Snapin" related code
