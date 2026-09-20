Attribute VB_Name = "PixelShifter"
Option Explicit

Function PixelShiftPages(ix As Integer) As Long
    ' deprecated
    Dim sHex As String: sHex = "&HA" & ix + 2
    PixelShiftPages = CLng(sHex)
End Function
