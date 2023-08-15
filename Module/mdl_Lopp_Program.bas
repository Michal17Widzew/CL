Attribute VB_Name = "mdl_Lopp_Program"
Option Explicit

Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long

Private Const THREAD_SUSPEND_RESUME As Long = &H2
Private Const THREAD_CREATE_SUSPENDED As Long = &H1

Dim threadID As Long

Public Sub StartContinuouslySaveThread(interval As Long)
    Dim threadHandle As Long
    threadHandle = CreateThread(ByVal 0&, ByVal 0&, AddressOf ContinuouslySaveData, ByVal interval, THREAD_CREATE_SUSPENDED, threadID)
    If threadHandle <> 0 Then
        ResumeThread threadHandle
    End If
End Sub

Private Sub ContinuouslySaveData(ByVal interval As Long)
    Do
        SaveModifiedData  ' Save the modified data
        Sleep interval ' Delay for the specified interval
    Loop
End Sub

