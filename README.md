<div align="center">

## WinPopupSimulator


</div>

### Description

My Code simulate a WinPoPup, i can Send message and receive message, i make a ocx( source code here) and another program for test this ocx(source code here)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-06-20 15:41:28
**By**             |[Serge Boivin \(Dark\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/serge-boivin-dark.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[WinPopupSi970066202002\.zip](https://github.com/Planet-Source-Code/serge-boivin-dark-winpopupsimulator__1-3223/archive/master.zip)

### API Declarations

```
Private Declare Function CloseHandle Lib "kernel32" (ByVal hHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFileName As Long, ByVal lpBuff As Any, ByVal nNrBytesToWrite As Long, lpNrOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwAccess As Long, ByVal dwShare As Long, ByVal lpSecurityAttrib As Long, ByVal dwCreationDisp As Long, ByVal dwAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateMailslot Lib "kernel32.dll" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
```





