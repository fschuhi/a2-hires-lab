Attribute VB_Name = "PixelShifter"
Option Explicit

Function PixelShiftPages(ix As Integer) As Long
    Dim sHex As String: sHex = "&HA" & ix + 2
    Debug.Print sHex
    PixelShiftPages = CLng(sHex)
End Function
