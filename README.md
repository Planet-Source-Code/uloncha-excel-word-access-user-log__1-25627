<div align="center">

## Excel/Word/Access user log


</div>

### Description

It keeps a log (IP, TIME (local), NetworkUserName) of those that open the office document you put the code in.

Example: Peter 7/29/2001 11:27:12 AM 172.19.20.22. Great to see who is opening your files.
 
### More Info
 
UserId, IP (taken from registry), Time stamp.

Read side effects. IP is taken from registry so it is necessary to point to its key.

I took some code from other programmer's cotributions, if you recognize the code please let me know and i will mention you.

Writes inputs in builtindocumentproperties.

if file is opened too oftenly and the log is not cleared, "comment" may become too big (i bet it will crash). Also, note that comments can be read by anyone and cleared too (unless document is write protected with password)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[uloncha](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/uloncha.md)
**Level**          |Beginner
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/uloncha-excel-word-access-user-log__1-25627/archive/master.zip)

### API Declarations

```
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_SZ = 1 ' Unicode
Public Const REG_DWORD = 4 ' 32-bit
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpsubkey As String, phkresult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpvaluename As String, ByVal lpreserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
```


### Source Code

```
Public Function getstring(hkey As Long, strpath As String, strvalue As String)
Dim keyhand, datatype, lResult, lDataBufSize As Long
Dim strBuf As String
Dim intZeroPos As Integer
 r = RegOpenKey(hkey, strpath, keyhand)
 lResult = RegQueryValueEx(keyhand, strvalue, 0&, lValueType, ByVal 0&, lDataBufSize)
 If lValueType = REG_SZ Then
 strBuf = String(lDataBufSize, " ")
 lResult = RegQueryValueEx(keyhand, strvalue, 0&, 0&, ByVal strBuf, lDataBufSize)
 If lResult = ERROR_SUCCESS Then
  intZeroPos = InStr(strBuf, Chr$(0))
  If intZeroPos > 0 Then
  getstring = Left$(strBuf, intZeroPos - 1)
  Else
  getstring = strBuf
  End If
 End If
 End If
End Function
Public Function NetworkUserName() As String
 Dim lpBuff As String * 25
 Dim retval As Long
 retval = GetUserName(lpBuff, 25)
 NetworkUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function
Public Function WorkstationID() As String
 Dim sBuffer As String * 255
 If GetComputerNameA(sBuffer, 255&) > 0 Then
 WorkstationID = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
 Else
 WorkstationID = "?"
 End If
End Function
Sub AUTO_Open()'put it in workbook_open in excel
ip = getstring(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\NetTrans\0001", "ipaddress")
ActiveWorkbook.BuiltinDocumentProperties(5) = ActiveWorkbook.BuiltinDocumentProperties(5) + vbCr + NetworkUserName & " " & Now & " " & ip
End Sub
```

