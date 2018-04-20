# Get-WsusReport
It's PowerShell cmdlet which can generates a report from corporate WSUS server about needed updates count of it's host groups.
It can show the report in format for using with zabbix_sender utility too.  
If you have any remark, contact me with Telegram @asand3r.  
Current stable verson:  
<b>0.2.3</b>  

## Main features
- [x] Show host group statistic from WSUS server.  
- [x] Save getting statistic to file.

## TODO List
- [ ] Make more suitable output format.    

## Installation
Please, read the relevant Wiki page.  

## Usage of current stable
You can use Get-Help cmdlet to view buit-in documentation, but here is some examples.

1. Gets report about Exchange hosts group:
```powershell
PS C:\> Get-WsusReport -Targetgroup Exchange

Name           : lon-cas1
IPAddress      : 192.168.1.1
OSVersion      : Windows Server 2008 R2 Standard Edition
LastSyncResult : Succeeded
NeedCount      : 2

Name           : lon-mbx1
IPAddress      : 192.168.1.2
OSVersion      : Windows Server 2008 R2 Enterprise Edition
LastSyncResult : Succeeded
NeedCount      : 5
```

2. Same in format for zabbix_sender utility:
```powershell
PS C:\> Get-WsusReport -Targetgroup Exchange -ZabbixSenderFormat
"LON-CAS1" os.updates.count 2
"LON-MBX1" os.updates.count 2
```
You may redefine item (os.updates.count) name using ZabbixItemName parameter.  

3. You can set connection properties in parameters:
```powershell
PS C:\> Get-WsusReport -Targetgroup Exchange -Server lon-wsus.domain.local -PortNumber 8530

Name           : lon-cas1
IPAddress      : 192.168.1.1
OSVersion      : Windows Server 2008 R2 Standard Edition
LastSyncResult : Succeeded
NeedCount      : 2

Name           : lon-mbx1
IPAddress      : 192.168.1.2
OSVersion      : Windows Server 2008 R2 Enterprise Edition
LastSyncResult : Succeeded
NeedCount      : 5
```

## Known issues