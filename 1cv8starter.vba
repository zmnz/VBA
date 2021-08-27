Const Above_Normal = 32768

Dim objSWbemObjectEx
Dim lngProcessID
Dim colProcesses
Dim setCountProc
Dim CountProc
Dim bUser
Dim strNameOfUser
Dim UserName
UserName = CreateObject("WScript.Network").UserName
setCountProc = 2 'задаем здесь
CountProc = 0
With WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer(".", "root\cimv2")
Set colProcesses = .ExecQuery("SELECT * FROM Win32_Process WHERE Name LIKE '1cv%'")
Set objSWbemObjectEx = .Get("Win32_ProcessStartup")
bUser = False

For Each objProcess in colProcesses
Return = objProcess.GetOwner(strNameOfUser)

If Return = 0 Then
If strNameOfUser = UserName Then
CountProc = CountProc + 1
'WScript.Echo setCountProc
If CountProc = setCountProc Then 
bUser = True
WScript.Echo "Запуск более "+CStr(setCountProc)+" экземпляров невозможен!"
End If
End If
End If
Next

If Not bUser Then
objSWbemObjectEx.PriorityClass = Above_Normal

' Create method of the Win32_Process class (Windows) (http://msdn.microsoft.com/en-us/libr...(v=vs.85).aspx)
If .Get("Win32_Process").Create( _
"""C:\Program Files (x86)\1cv8\8.3.18.1289\bin\1cv8.exe""", _
"C:\Program Files (x86)\1cv8\8.3.18.1289\bin", _
objSWbemObjectEx, _
lngProcessID _
) <> 0 Then
WScript.Echo "Can't start process ""%Program Files (x86)\1cv8\8.3.18.1289\bin""."
End If

Set objSWbemObjectEx = Nothing
End If
End With

WScript.Quit 0
