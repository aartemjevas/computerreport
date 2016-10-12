# computerreport
Gets services and installed applications from remote computers
```powershell
<#
.SYNOPSIS
Gets application and services list from remote computers

.DESCRIPTION
Gets application and services list from remote computers

.PARAMETER ComputerList
Path to text file with computernames. One computername per line.

.PARAMETER Credentials
Credentials for connecting to remote computers

.PARAMETER OutputPath
PAth where reports will be saved

.PARAMETER ReportOwner
Person who created reports

.PARAMETER LogName
Log file name

.EXAMPLE
.\computerReport.ps1 -ComputerList .\pcs.txt -OutputPath .\Reports -LogName 'pclog.log' -Credentials (Get-Credential)


#>
```
### Index Page

![alt tag](http://i.imgur.com/wdOnj3F.png)

### Computer Page

![alt tag](http://i.imgur.com/CWCvatu.png)

