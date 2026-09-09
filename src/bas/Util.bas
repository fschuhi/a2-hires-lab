Attribute VB_Name = "Util"
Option Explicit

' Util -- hex/bit conversion helpers for a2-hires-lab.
' Naming: Hungarian notation (iX Integer, lX Long, sX String).
' No cl prefix here -- this is a standard module, not a class.

Public Function ByteToHex(ByVal lVal As Long) As String
    ' 0-255 -> "00"-"FF"
    ByteToHex = Right$("0" & Hex$(lVal), 2)
End Function

Public Function HexToByte(ByVal sHex As String) As Long
    ' "00"-"FF" -> 0-255
    HexToByte = CLng("&H" & sHex)
End Function

Public Function GetBit(ByVal lVal As Long, ByVal iBit As Long) As Long
    ' Extract one bit (0 or 1) at position iBit (0 = LSB).
    GetBit = (lVal \ (2 ^ iBit)) Mod 2
End Function

Public Function SetBit(ByVal lVal As Long, ByVal iBit As Long, ByVal iB As Long) As Long
    ' Set (iB <> 0) or clear (iB = 0) one bit at position iBit.
    Dim lMask As Long
    lMask = 2 ^ iBit
    If iB <> 0 Then
        SetBit = lVal Or lMask
    Else
        SetBit = lVal And (Not lMask)
    End If
End Function

Public Function ByteToBits(ByVal lVal As Long) As String
    ' 0-127 -> "0000000"-"1111111", bit 6 leftmost, bit 0 rightmost.
    ' This is the chapter's "Bits" column storage order, not pixel order.
    Dim sBits As String
    Dim iBit As Long
    sBits = ""
    For iBit = 6 To 0 Step -1
        sBits = sBits & GetBit(lVal, iBit)
    Next iBit
    ByteToBits = sBits
End Function

Public Function BitsToByte(ByVal sBits As String) As Long
    ' Inverse of ByteToBits.
    Dim lVal As Long
    Dim iPos As Long
    Dim iBit As Long
    lVal = 0
    For iPos = 1 To Len(sBits)
        iBit = 7 - iPos   ' char 1 -> bit 6, char 7 -> bit 0
        If Mid$(sBits, iPos, 1) = "1" Then
            lVal = SetBit(lVal, iBit, 1)
        End If
    Next iPos
    BitsToByte = lVal
End Function

Public Function HexToBits(ByVal sHex As String) As String
    ' "00"-"FF" -> 7-bit binary string, high bit masked off.
    ' Callable directly from a worksheet cell, e.g. =HexToBits(R4)
    HexToBits = ByteToBits(HexToByte(sHex) And &H7F)
End Function
