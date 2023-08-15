Attribute VB_Name = "mdl_Memory_Hook"
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Dim f1holder As Integer
Dim timer_pos As Long
'API Declaration
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As Long
Public Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Public Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Function WriteAByte(gamewindowtext As String, address As Long, value As Byte)
Dim hWnd As Long
Dim pid As Long
Dim phandle As Long
hWnd = FindWindow(vbNullString, gamewindowtext)
If (hWnd = 0) Then
MsgBox "The Game Is Not Working", vbCritical, "Error"
End
Exit Function
End If
GetWindowThreadProcessId hWnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
WriteProcessMemory phandle, address, value, 1, 0&
CloseHandle hProcess
End Function
Public Function WriteAnInt(gamewindowtext As String, address As Long, value As Integer)
Dim hWnd As Long
Dim pid As Long
Dim phandle As Long
hWnd = FindWindow(vbNullString, gamewindowtext)
If (hWnd = 0) Then
MsgBox "The Game Is Not Working", vbCritical, "Error"
End
End If
GetWindowThreadProcessId hWnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
WriteProcessMemory phandle, address, value, 2, 0&
CloseHandle hProcess
End Function
Public Function WriteALong(gamewindowtext As String, address As Long, value As Long)
Dim hWnd As Long
Dim pid As Long
Dim phandle As Long
hWnd = FindWindow(vbNullString, gamewindowtext)
If (hWnd = 0) Then
MsgBox "The Game Is Not Working", vbCritical, "Error"
End
Exit Function
End If
GetWindowThreadProcessId hWnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
WriteProcessMemory phandle, address, value, 4, 0&
CloseHandle hProcess
End Function
Public Function ReadAByte(gamewindowtext As String, address As Long, valbuffer As Byte)
Dim hWnd As Long
Dim pid As Long
Dim phandle As Long
hWnd = FindWindow(vbNullString, gamewindowtext)
If (hWnd = 0) Then
MsgBox "The Game Is Not Working", vbCritical, "Error"
End
Exit Function
End If
GetWindowThreadProcessId hWnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
ReadProcessMem phandle, address, valbuffer, 1, 0&
CloseHandle hProcess
End Function
Public Function ReadAnInt(gamewindowtext As String, address As Long, valbuffer As Integer)
Dim hWnd As Long
Dim pid As Long
Dim phandle As Long
hWnd = FindWindow(vbNullString, gamewindowtext)
If (hWnd = 0) Then
MsgBox "The Game Is Not Working", vbCritical, "Error"
End
Exit Function
End If
GetWindowThreadProcessId hWnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
ReadProcessMem phandle, address, valbuffer, 2, 0&
CloseHandle hProcess
End Function
Public Function ReadALong(gamewindowtext As String, address As Long, valbuffer As Long)
Dim hWnd As Long
Dim pid As Long
Dim phandle As Long
hWnd = FindWindow(vbNullString, gamewindowtext)
If (hWnd = 0) Then
MsgBox "The Game Is Not Working", vbCritical, "Error"
End
Exit Function
End If
GetWindowThreadProcessId hWnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
ReadProcessMem phandle, address, valbuffer, 4, 0&
CloseHandle hProcess
End Function
Public Function ReadAFloat(gamewindowtext As String, address As Long, valbuffer As Single)
Dim hWnd As Long
Dim pid As Long
Dim phandle As Long
hWnd = FindWindow(vbNullString, gamewindowtext)
If (hWnd = 0) Then
MsgBox "The Game Is Not Working", vbCritical, "Error"
End
Exit Function
End If
GetWindowThreadProcessId hWnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
ReadProcessMem phandle, address, valbuffer, 4, 0&
CloseHandle hProcess
End Function
 
Public Function WriteAFloat(gamewindowtext As String, address As Long, value As Single)
Dim hWnd As Long
Dim pid As Long
Dim phandle As Long
hWnd = FindWindow(vbNullString, gamewindowtext)
If (hWnd = 0) Then
MsgBox "The Game Is Not Working", vbCritical, "Error"
End
Exit Function
End If
GetWindowThreadProcessId hWnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
MsgBox "Can't get ProcessId", vbCritical, "Error"
Exit Function
End If
WriteProcessMemory phandle, address, value, 4, 0&
CloseHandle hProcess
End Function


