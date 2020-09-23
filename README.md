<div align="center">

## CPU Usage on all 32bit Windows OS


</div>

### Description

A simple to use class that returns the current CPU load usage of the system as a percentage. Works with all 32bit Windows operating systems (9x, ME, NT, 2000, XP) and also works with multiple processors.

I couldn't find any CPU usage code on PSC that would work with XP, so I had a hunt around the web and came across the NtQuerySystemInformation API.

The class detects the OS and uses the appropriate CPU usage retrieval system.
 
### More Info
 
Developers should find this class straight-forward to use. Let me know if you have any dfficulties.

CurrentCPUUsage As Long (percentage of current CPU loading)


<span>             |<span>
---                |---
**Submitted On**   |2003-04-18 22:53:04
**By**             |[ketamine](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ketamine.md)
**Level**          |Advanced
**User Rating**    |4.1 (29 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CPU\_Usage\_1575974182003\.zip](https://github.com/Planet-Source-Code/ketamine-cpu-usage-on-all-32bit-windows-os__1-44847/archive/master.zip)

### API Declarations

```
Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, ByVal lpStructure As Long, ByVal dwSize As Long, ByVal dwReserved As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
```





