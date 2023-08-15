Attribute VB_Name = "Module1"
Dim Groupe_Mod As Byte
Dim Team As String

Public Function ChoiceTeamCfg(txtGroupe As TextBox, txtTeam As TextBox)

    Dim Change_Team_Adrr As Long 'offset choice team
    Dim iniFilePath As String
    Dim Grupa As String

    
    iniFilePath = App.Path & "\Tournament.cfg"
    
    
    ' load mdl_system ini module
    '                           |
    '                           |
    '                           |
    '                           V
 
       Grupa = sGetINI(iniFilePath, "Select Team", "Groupe", "0")
       Team = sGetINI(iniFilePath, "Select Team", "Team", "0")

 ' check and convert string ini data to value
Select Case Grupa
    Case "A"
        Groupe_Mod = 1
    Case "B"
        Groupe_Mod = 2
    Case "C"
        Groupe_Mod = 3
    Case "D"
        Groupe_Mod = 4
    Case "E"
        Groupe_Mod = 5
    Case "F"
        Groupe_Mod = 6
    Case "G"
        Groupe_Mod = 7
    Case "H"
        Groupe_Mod = 8
End Select


txtGroupe = Groupe_Mod
txtTeam = Team
     'Change_Team_Adrr = &H3A7B080 ' Wybierz zespol do turnieju 0 = Off 1 = On
     'Call ReadAByte("Pro Evolution Soccer 6", Change_Team_Adrr, Choice)
'MsgBox Choice
End Function

Public Sub SaveModifiedData()
    Dim Choose_Adrr As Long
    
    Choose_Adrr = &H3A7B080 ' FIRST TEAM'
    'A=1'B=2'C=3'D=4'E=5'F=6'G=7'H=8
    
'<Magic key JUMP Pes 6 array tabele> ;)
    On Error Resume Next ' Enable error handling
'Grupa A
If Groupe_Mod = 1 And Team = 1 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr, 1)
If Groupe_Mod = 1 And Team = 2 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 32, 1)
If Groupe_Mod = 1 And Team = 3 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 64, 1)
If Groupe_Mod = 1 And Team = 4 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 96, 1)
'Grupa B
If Groupe_Mod = 2 And Team = 1 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 8, 1)
If Groupe_Mod = 2 And Team = 2 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 40, 1)
If Groupe_Mod = 2 And Team = 3 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 72, 1)
If Groupe_Mod = 2 And Team = 4 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 104, 1)
'Grupa C
If Groupe_Mod = 3 And Team = 1 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 16, 1)
If Groupe_Mod = 3 And Team = 2 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 48, 1)
If Groupe_Mod = 3 And Team = 3 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 80, 1)
If Groupe_Mod = 3 And Team = 4 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 112, 1)
'Grupa D
If Groupe_Mod = 4 And Team = 1 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 24, 1)
If Groupe_Mod = 4 And Team = 2 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 56, 1)
If Groupe_Mod = 4 And Team = 3 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 88, 1)
If Groupe_Mod = 4 And Team = 4 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 120, 1)
'Grupa E
If Groupe_Mod = 5 And Team = 1 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 128, 1)
If Groupe_Mod = 5 And Team = 2 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 160, 1)
If Groupe_Mod = 5 And Team = 3 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 192, 1)
If Groupe_Mod = 5 And Team = 4 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 224, 1)
'Grupa F
If Groupe_Mod = 6 And Team = 1 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 136, 1)
If Groupe_Mod = 6 And Team = 2 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 168, 1)
If Groupe_Mod = 6 And Team = 3 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 200, 1)
If Groupe_Mod = 6 And Team = 4 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 232, 1)
'Grupa G
If Groupe_Mod = 7 And Team = 1 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 144, 1)
If Groupe_Mod = 7 And Team = 2 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 176, 1)
If Groupe_Mod = 7 And Team = 3 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 208, 1)
If Groupe_Mod = 7 And Team = 4 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 240, 1)
'Grupa H
If Groupe_Mod = 8 And Team = 1 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 152, 1)
If Groupe_Mod = 8 And Team = 2 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 184, 1)
If Groupe_Mod = 8 And Team = 3 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 216, 1)
If Groupe_Mod = 8 And Team = 4 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 248, 1)
'Grupa TEST
'If Groupe_Mod = 6 And Team = 2 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 168, 1)
        'Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 80, 1)

    On Error GoTo 0 ' Disable error handling
End Sub

