VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Champion_League"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Team() As Integer ' Team array memory CL Cup in PES6
Dim DataSztos As Byte ' counter euro cup = dataCounterOffset

Dim TeamOffset As Long
Dim dataCounterOffset As Long ' Button Ok in-game sztos = 16 32 team /2 =16


Public Function EuroRelase(lstBox As ListBox)
    Dim i As Long
    Dim t As Long


    TeamOffset = &H3A7B07C ' Team ID location offset in memory for PES6

    t = TeamOffset
    lstBox.Clear ' Clear the ListBox control

    ReDim Team(0 To 33) As Integer ' Resize the Team array

    On Error Resume Next ' Enable error handling

    For i = 1 To UBound(Team) - 1
        Call ReadAnInt("Pro Evolution Soccer 6", t, Team(i))
        lstBox.AddItem Team(i)
        t = t + 8
    Next i

    On Error GoTo 0 ' Disable error handling
End Function

Public Sub ModifyTeamData(Index As Integer, value As Integer)
    If Index >= 1 And Index <= UBound(Team) Then
        Team(Index) = value
    End If
End Sub

Public Sub SaveModifiedData()
    Dim t As Long
    Dim Button_ok As Long
    
    dataCounterOffset = &H3A79FD2
    Button_ok = &H20
   
    t = TeamOffset

    On Error Resume Next ' Enable error handling

    For i = 1 To UBound(Team) - 1
        Call WriteAnInt("Pro Evolution Soccer 6", t, Team(i))
        t = t + 8
    Next i
    Call WriteALong("Pro Evolution Soccer 6", dataCounterOffset, Button_ok) ' save unlock button Accept In Pes6 Tournament>All Cup ? maybe league
   
    On Error GoTo 0 ' Disable error handling
End Sub




