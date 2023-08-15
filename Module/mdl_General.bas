Attribute VB_Name = "mdl_General"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1


Public Sub Run_Game()
Dim appPath As String
Dim params As String
    Dim iniFilePath As String
    Dim m_data As String
    
    iniFilePath = App.Path & "\Tournament.cfg"
    
 mdata = sGetINI(iniFilePath, "Run As PES6", "Game", "0")
appPath = mdata 'App.Path & "\Pes6.exe" ' Œcie¿ka do aplikacji, któr¹ chcesz uruchomiæ
params = "" ' Parametry aplikacji (opcjonalne)

ShellExecute frmMain.hWnd, "runas", appPath, params, vbNullString, vbNormalFocus
End Sub
