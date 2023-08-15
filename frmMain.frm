VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CL_Team_Server"
   ClientHeight    =   4620
   ClientLeft      =   7245
   ClientTop       =   1620
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   308
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   643
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   5400
      TabIndex        =   6
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   5400
      TabIndex        =   5
      Top             =   4200
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   6300
      Left            =   2880
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5880
      Top             =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   5280
      TabIndex        =   1
      Top             =   5280
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   6300
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Champion League Teams Server by Micha³17Widzew"
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   3825
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mem Data Index 1 to 32"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tournamnet.cfg"
      Height          =   195
      Left            =   3000
      TabIndex        =   3
      Top             =   1560
      Width           =   1065
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CL_Team As New cls_Champion_League ' classs load memory data Pes Champion Legaue Mode
Dim Choose_Team As New cls_Choose_Team

Dim newArray() As Integer ' data cfg convert to byte and send to save

Private Sub Command1_Click() 'test button
CreateArrayFromListBox List2
End Sub

Private Sub Form_Load()
Run_Game  ' runas  pes6.exe
  LoadIniFile List2   ' <Load Tournaments.cfg>
  
CreateArrayFromListBox List2 '<new collect from listbox to Array>'1 zamien stringi z cfg na value i zapisz w pamieci
Choose_Team.ChoiceTeamCfg txt(0), txt(1) ' load cfg'2 ustal jakim zepolem bedziesz gral w grze to juz bedzie niedostepne
   CL_Team.EuroRelase List1 ' relase null database'3

   
End Sub


Private Sub Timer1_Timer()
    Dim teamData(0 To 31) As Integer
    Dim i As Integer
                                            '<Grupa A>
CL_Team.ModifyTeamData 1, newArray(1)        '1
CL_Team.ModifyTeamData 5, newArray(2)       '2
CL_Team.ModifyTeamData 9, newArray(3)       '3
CL_Team.ModifyTeamData 13, newArray(4)      '4
    '                                       <Grupa B>
CL_Team.ModifyTeamData 2, newArray(5)       '1
CL_Team.ModifyTeamData 6, newArray(6)       '2
CL_Team.ModifyTeamData 10, newArray(7)      '3
CL_Team.ModifyTeamData 14, newArray(8)      '4
                                            '<Grupa C>
CL_Team.ModifyTeamData 3, newArray(9)       '1
CL_Team.ModifyTeamData 7, newArray(10)      '2
CL_Team.ModifyTeamData 11, newArray(11)     '3
CL_Team.ModifyTeamData 15, newArray(12)     '4
                                            '<Grupa D>
CL_Team.ModifyTeamData 4, newArray(13)      '1
CL_Team.ModifyTeamData 8, newArray(14)      '2
CL_Team.ModifyTeamData 12, newArray(15)     '3
CL_Team.ModifyTeamData 16, newArray(16)     '4
                                            '<Grupa E>
CL_Team.ModifyTeamData 17, newArray(17)     '1
CL_Team.ModifyTeamData 21, newArray(18)     '2
CL_Team.ModifyTeamData 25, newArray(19)     '3
CL_Team.ModifyTeamData 29, newArray(20)     '4
                                            '<Grupa F>
CL_Team.ModifyTeamData 18, newArray(21)     '1
CL_Team.ModifyTeamData 22, newArray(22)     '2
CL_Team.ModifyTeamData 26, newArray(23)     '3
CL_Team.ModifyTeamData 30, newArray(24)     '4
                                            '<Grupa G>
CL_Team.ModifyTeamData 19, newArray(25)     '1
CL_Team.ModifyTeamData 23, newArray(26)     '2
CL_Team.ModifyTeamData 27, newArray(27)     '3
CL_Team.ModifyTeamData 31, newArray(28)     '4
                                            '<Grupa H>
CL_Team.ModifyTeamData 20, newArray(29)     '1
CL_Team.ModifyTeamData 24, newArray(30)     '2
CL_Team.ModifyTeamData 28, newArray(31)     '3
CL_Team.ModifyTeamData 32, newArray(32)     '4


    'For i = 1 To 32
        'teamData(i) = i - 1
       ' CL_Team.ModifyTeamData i, teamData(i)
    'Next i

    CL_Team.SaveModifiedData ' save new groupe /change in memory Pes6
    Choose_Team.SaveModifiedData 'zapisz wybrany przez ciebie zespol/ your team
   ' MsgBox newArray(1)
End Sub

'//  Nowa kolekcja z listboxa ¿eby ja przekonwartowac na integer i to pojdzie do timera i zostanie zapisania w pamieci(new array from listbox2)

Public Function CreateArrayFromListBox(lstBox As ListBox) As Integer()

    ReDim newArray(1 To lstBox.ListCount) As Integer

    Dim i As Integer
    For i = 1 To lstBox.ListCount
        newArray(i) = CInt(lstBox.List(i - 1))
    Next i

    CreateArrayFromListBox = newArray
    'lstBox.Clear

End Function

