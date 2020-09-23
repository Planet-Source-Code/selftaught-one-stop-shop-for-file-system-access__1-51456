Attribute VB_Name = "mMain"
Option Explicit


Public Sub Main()
    Dim loTemp As New cTest
    loTemp.Test
End Sub

Public Function MakeQWord(ByVal piLow As Long, ByVal piHigh As Long) As Double
    Const MAX_DWORD = &HFFFF
    MakeQWord = (piHigh * MAX_DWORD + piLow)
End Function
