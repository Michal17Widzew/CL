Attribute VB_Name = "mdl_System_ini_Config"
Rem API DECLARATIONS
'//   ProfileString from ApiGudie
Rem Rest code ist  me

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
                 "GetPrivateProfileStringA" (ByVal lpApplicationName _
                 As String, ByVal lpKeyName As Any, ByVal lpDefault _
                 As String, ByVal lpReturnedString As String, ByVal _
                 nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias _
                 "WritePrivateProfileStringA" (ByVal lpApplicationName _
                 As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
                 ByVal lpFileName As String) As Long
                 
Public GArray(1 To 32) As Byte
                 
                 
Public Function sGetINI(sINIFile As String, sSection As String, sKey _
                As String, sDefault As String) As String
    Dim sTemp As String * 256
    Dim nLength As Integer
    sTemp = Space$(256)
    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, _
              255, sINIFile)
    sGetINI = Left$(sTemp, nLength)
End Function
Public Sub writeINI(sINIFile As String, sSection As String, sKey _
           As String, sValue As String)
    Dim n As Integer
    Dim sTemp As String
    sTemp = sValue
    Rem Replace any CR/LF characters with spaces
    For n = 1 To Len(sValue)
        If Mid$(sValue, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf _
        Then Mid$(sValue, n) = " "
    Next n
    n = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)
End Sub



Public Function LoadIniFile(Optional lstBox As ListBox)
    Dim i As Byte
    Dim iniFilePath As String
    iniFilePath = App.Path & "\Tournament.cfg"
    
    ' Grupa A from Ini
    For i = 1 To 4
        intValue = Val(sGetINI(iniFilePath, "Groupe A", "Team_" & i, "0"))
        If intValue >= 0 And intValue <= 255 Then
            GArray(i) = CInt(intValue)
        Else
            GArray(i) = 0
        End If
    Next i

    ' Grupa B from Ini
    For i = 1 To 4
        intValue = Val(sGetINI(iniFilePath, "Groupe B", "Team_" & i, "0"))
        If intValue >= 0 And intValue <= 255 Then
            GArray(i + 4) = CInt(intValue)
        Else
            GArray(i + 4) = 0
        End If
    Next i

    ' Grupa C from Ini
    For i = 1 To 4
        intValue = Val(sGetINI(iniFilePath, "Groupe C", "Team_" & i, "0"))
        If intValue >= 0 And intValue <= 255 Then
            GArray(i + 8) = CInt(intValue)
        Else
            GArray(i + 8) = 0
        End If
    Next i

    ' Grupa D from Ini
    For i = 1 To 4
        intValue = Val(sGetINI(iniFilePath, "Groupe D", "Team_" & i, "0"))
        If intValue >= 0 And intValue <= 255 Then
            GArray(i + 12) = CInt(intValue)
        Else
            GArray(i + 12) = 0
        End If
    Next i

    ' Grupa E from Ini
    For i = 1 To 4
        intValue = Val(sGetINI(iniFilePath, "Groupe E", "Team_" & i, "0"))
        If intValue >= 0 And intValue <= 255 Then
            GArray(i + 16) = CInt(intValue)
        Else
            GArray(i + 16) = 0
        End If
    Next i

    ' Grupa F from Ini
    For i = 1 To 4
        intValue = Val(sGetINI(iniFilePath, "Groupe F", "Team_" & i, "0"))
        If intValue >= 0 And intValue <= 255 Then
            GArray(i + 20) = CInt(intValue)
        Else
            GArray(i + 20) = 0
        End If
    Next i

    ' Grupa G from Ini
    For i = 1 To 4
        intValue = Val(sGetINI(iniFilePath, "Groupe G", "Team_" & i, "0"))
        If intValue >= 0 And intValue <= 255 Then
            GArray(i + 24) = CInt(intValue)
        Else
            GArray(i + 24) = 0
        End If
    Next i

    ' Grupa H from Ini
    For i = 1 To 4
        intValue = Val(sGetINI(iniFilePath, "Groupe H", "Team_" & i, "0"))
        If intValue >= 0 And intValue <= 255 Then
            GArray(i + 28) = CInt(intValue)
        Else
            GArray(i + 28) = 0
        End If
    Next i

    If Not lstBox Is Nothing Then
        Dim j As Integer
        For j = 1 To 32
            lstBox.AddItem GArray(j)
        Next j
    End If
End Function


