VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_LoadTournament_CFGdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 'Load File my Format File by Micha³17Widzew 12.07.2023
 'filePath = App.Path & "\Tournament.ini" ' Zmieñ na w³aœciw¹ œcie¿kê do pliku
Public Function ReadDataFromFile(filePath As String, lstbox As ListBox)
    'Dim filePath As String
    Dim fileNumber As Integer
    Dim line As String
    Dim groupName As String
    Dim element1 As Long
    Dim element2 As Long
    Dim element3 As Long
    Dim element4 As Long
    Dim dataArray(1 To 32) As Long ' Tablica przechowuj¹ca dane z grup¹
    
   ' filePath = App.Path & "\Tournament.ini" ' Zmieñ na w³aœciw¹ œcie¿kê do pliku
    
    ' Otwieranie pliku do odczytu
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    
    ' Przetwarzanie pliku linia po linii
    Do Until EOF(fileNumber)
        Line Input #fileNumber, line
        
        ' Sprawdzanie wzorca dla pocz¹tku grupy A
        If Left$(Trim$(line), 7) = "<GroupA" Then
            ' Wyodrêbnianie nazwy grupy
            groupName = "A"
        End If
        
        ' Sprawdzanie wzorca dla elementów grupy A
        If Left$(Trim$(line), 10) = "<element1>" Then
            ' Wyodrêbnianie wartoœci elementu 1
            element1 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        ElseIf Left$(Trim$(line), 10) = "<element2>" Then
            ' Wyodrêbnianie wartoœci elementu 2
            element2 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        ElseIf Left$(Trim$(line), 10) = "<element3>" Then
            ' Wyodrêbnianie wartoœci elementu 3
            element3 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        ElseIf Left$(Trim$(line), 10) = "<element4>" Then
            ' Wyodrêbnianie wartoœci elementu 4
            element4 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        End If
        
        ' Sprawdzanie wzorca dla koñca grupy B
        If Trim$(line) = "</GroupA>" Then
            ' Dodawanie danych do tablicy
            dataArray(1) = element1
            dataArray(2) = element2
            dataArray(3) = element3
            dataArray(4) = element4
            
            
            ' Resetowanie zmiennych
            groupName = ""
            element1 = 0
            element2 = 0
            element3 = 0
            element4 = 0
        End If
        
        ' Sprawdzanie wzorca dla pocz¹tku grupy B
        If Left$(Trim$(line), 7) = "<GroupB" Then
            ' Wyodrêbnianie nazwy grupy
            groupName = "B"
        End If
        
     ' Sprawdzanie wzorca dla elementów grupy B
        If Left$(Trim$(line), 10) = "<element1>" Then
            ' Wyodrêbnianie wartoœci elementu 1
            element1 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        ElseIf Left$(Trim$(line), 10) = "<element2>" Then
            ' Wyodrêbnianie wartoœci elementu 2
            element2 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        ElseIf Left$(Trim$(line), 10) = "<element3>" Then
            ' Wyodrêbnianie wartoœci elementu 3
            element3 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        ElseIf Left$(Trim$(line), 10) = "<element4>" Then
            ' Wyodrêbnianie wartoœci elementu 4
            element4 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        End If
        
        ' Sprawdzanie wzorca dla koñca grupy B
        If Trim$(line) = "</GroupB>" Then
            ' Dodawanie danych do tablicy
            dataArray(5) = element1
            dataArray(6) = element2
            dataArray(7) = element3
            dataArray(8) = element4
            
            ' Wyœwietlanie lub przetwarzanie danych dla grupy B
            Debug.Print "Grupa: " & groupName
            For i = 1 To 4
                Debug.Print "element" & i & ": " & dataArray(i)
            Next i
            Debug.Print
            
            ' Resetowanie zmiennych
            groupName = ""
            element1 = 0
            element2 = 0
            element3 = 0
            element4 = 0
        End If
        
                ' Sprawdzanie wzorca dla pocz¹tku grupy C
        If Left$(Trim$(line), 7) = "<GroupC" Then
            ' Wyodrêbnianie nazwy grupy
            groupName = "C"
        End If
        
     ' Sprawdzanie wzorca dla elementów grupy C
        If Left$(Trim$(line), 10) = "<element1>" Then
            ' Wyodrêbnianie wartoœci elementu 1
            element1 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        ElseIf Left$(Trim$(line), 10) = "<element2>" Then
            ' Wyodrêbnianie wartoœci elementu 2
            element2 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        ElseIf Left$(Trim$(line), 10) = "<element3>" Then
            ' Wyodrêbnianie wartoœci elementu 3
            element3 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        ElseIf Left$(Trim$(line), 10) = "<element4>" Then
            ' Wyodrêbnianie wartoœci elementu 4
            element4 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        End If
        
        ' Sprawdzanie wzorca dla koñca grupy C
        If Trim$(line) = "</GroupC>" Then
            ' Dodawanie danych do tablicy
            dataArray(9) = element1
            dataArray(10) = element2
            dataArray(11) = element3
            dataArray(12) = element4
            
            ' Wyœwietlanie lub przetwarzanie danych dla grupy C
            Debug.Print "Grupa: " & groupName
            For i = 1 To 4
                Debug.Print "element" & i & ": " & dataArray(11)
            Next i
            Debug.Print
            
            ' Resetowanie zmiennych
            groupName = ""
            element1 = 0
            element2 = 0
            element3 = 0
            element4 = 0
        End If
        
    ' Sprawdzanie wzorca dla pocz¹tku grupy D
        If Left$(Trim$(line), 7) = "<GroupD" Then
            ' Wyodrêbnianie nazwy grupy D
            groupName = "D"
        End If
        
     ' Sprawdzanie wzorca dla elementów grupy D
        If Left$(Trim$(line), 10) = "<element1>" Then
            ' Wyodrêbnianie wartoœci elementu 1
            element1 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        ElseIf Left$(Trim$(line), 10) = "<element2>" Then
            ' Wyodrêbnianie wartoœci elementu 2
            element2 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        ElseIf Left$(Trim$(line), 10) = "<element3>" Then
            ' Wyodrêbnianie wartoœci elementu 3
            element3 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        ElseIf Left$(Trim$(line), 10) = "<element4>" Then
            ' Wyodrêbnianie wartoœci elementu 4
            element4 = CLng(Mid$(line, InStr(line, ">") + 1, InStrRev(line, "<") - InStr(line, ">") - 1))
        End If
        
        ' Sprawdzanie wzorca dla koñca grupy D
        If Trim$(line) = "</GroupD>" Then
            ' Dodawanie danych do tablicy
            dataArray(13) = element1
            dataArray(14) = element2
            dataArray(15) = element3
            dataArray(16) = element4
            
            ' Wyœwietlanie lub przetwarzanie danych dla grupy D
            Debug.Print "Grupa: " & groupName
            For i = 1 To 4
                Debug.Print "element" & i & ": " & dataArray(i)
            Next i
            Debug.Print
            
            ' Resetowanie zmiennych
            groupName = ""
            element1 = 0
            element2 = 0
            element3 = 0
            element4 = 0
        End If
        ' Analogiczne sekcje kodu dla grup C i D
        
        ' ...
        
    Loop
    lstbox.AddItem "element" & i & ": " & dataArray(2)
    ' Zamykanie pliku
    Close fileNumber
End Function


