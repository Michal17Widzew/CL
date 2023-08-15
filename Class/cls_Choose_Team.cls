VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Choose_Team"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Groupe_Mod As Byte
Dim Team As String

Public Function ChoiceTeamCfg(Optional txtGroupe As TextBox, Optional txtTeam As TextBox)

    Dim Change_Team_Adrr As Long 'offset choice team
    Dim iniFilePath As String
    Dim Grupa As String

    
    iniFilePath = App.Path & "\Tournament.cfg"
    
    
    ' load mdl_system ini module
    '                           |
    '                           |
    '                           |
    '                           V
 
       Grupa = sGetINI(iniFilePath, "Select Team", "Groupe", "0")       '
       Team = sGetINI(iniFilePath, "Select Team", "Team", "0")          '
    '------------------------------------------------------------------ '
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

'debug option
txtGroupe = Groupe_Mod
txtTeam = Team
     'Change_Team_Adrr = &H3A7B080 ' Wybierz zespol do turnieju 0 = Off 1 = On
     'Call ReadAByte("Pro Evolution Soccer 6", Change_Team_Adrr, Choice)
'MsgBox Choice
End Function

Public Sub SaveModifiedData()
    Dim Choose_Adrr As Long
    
    Choose_Adrr = &H3A7B080 ' FIRST TEAM'
    'Data convert string ini to array A=1'B=2'C=3'D=4'E=5'F=6'G=7'H=8
    
'<Magic key JUMP Pes 6 array tabele> ;)
    On Error Resume Next ' Enable error handling
Select Case Groupe_Mod
    Case 1 ' Grupa A
        Select Case Team
            Case 1: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr, 1)
            Case 2: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H20, 1)
            Case 3: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H40, 1)
            Case 4: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H60, 1)
        End Select
    Case 2 ' Grupa B
        Select Case Team
            Case 1: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H8, 1)
            Case 2: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H28, 1)
            Case 3: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H48, 1)
            Case 4: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H68, 1)
        End Select
    Case 3 ' Grupa C
        Select Case Team
            Case 1: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H10, 1)
            Case 2: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H30, 1)
            Case 3: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H50, 1)
            Case 4: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H70, 1)
        End Select
    Case 4 ' Grupa D
        Select Case Team
            Case 1: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H18, 1)
            Case 2: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H38, 1)
            Case 3: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H58, 1)
            Case 4: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H78, 1)
        End Select
    Case 5 ' Grupa E
        Select Case Team
            Case 1: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H80, 1)
            Case 2: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &HA0, 1)
            Case 3: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &HC0, 1)
            Case 4: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &HE0, 1)
        End Select
    Case 6 ' Grupa F
        Select Case Team
            Case 1: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H88, 1)
            Case 2: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &HA8, 1)
            Case 3: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &HC8, 1)
            Case 4: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &HE8, 1)
        End Select
    Case 7 ' Grupa G
        Select Case Team
            Case 1: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H90, 1)
            Case 2: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &HB0, 1)
            Case 3: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &HD0, 1)
            Case 4: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &HF0, 1)
        End Select
    Case 8 ' Grupa H
        Select Case Team
            Case 1: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &H98, 1)
            Case 2: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &HB8, 1)
            Case 3: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &HD8, 1)
            Case 4: Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + &HF8, 1)
        End Select
End Select


'Grupa TEST
'If Groupe_Mod = 6 And Team = 2 Then Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 168, 1)
        'Call WriteAByte("Pro Evolution Soccer 6", Choose_Adrr + 80, 1)

    On Error GoTo 0 ' Disable error handling
End Sub


