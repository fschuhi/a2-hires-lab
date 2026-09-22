Attribute VB_Name = "ScreenMemory"
' Copyright (c) 2026 Frank Schuhardt
' SPDX-License-Identifier: MIT

Option Explicit

' ScreenMemory -- writes into the HGR page 1 memory model on 'Hires Memory'.
'
' 'Hires Memory' lists the 192 screen lines in address order (C3:C194),
' with the 40 bytes of each line in F:AS. 'Hires Pixels' and 'Hires Screen'
' derive everything else from those bytes, so writing here is all it takes
' to put something on the screen.
'
' Coordinates are 0-based, like the labels on 'Hires Screen' and the 6502:
' screen rows 0..191, pixel columns 0..279.
'
' First step only: this ignores the $FF boundary problem the game has with
' 280 columns, does no masking and does not clear what was there before.

Private Const MEMORY_SHEET As String = "Hires Memory"
Private Const MEMORY_ADDR_RANGE As String = "C3:C194"
Private Const MEMORY_FIRST_DATA_COL As Long = 6     ' column F = byte offset 0

Private Const HGR1_PAGE As Long = &H20
Private Const SCREEN_ROWS As Long = 192
Private Const SCREEN_COLS As Long = 280
Private Const BYTES_PER_LINE As Long = 40
Private Const PIXELS_PER_BYTE As Long = 7
Private Const SPRITE_ROWS As Long = 11
Private Const SPRITE_BYTES As Long = 3
Private Const OFFSET_TABLE_COLS As Long = 16         ' RowToOffset*Table is 12 x 16


Public Sub PlaceShiftedSprite(ByVal iRow As Integer, ByVal iCol As Integer)
    ' Places the sprite so that its top-left pixel lands at (iRow, iCol).
    ' At least that one pixel must be on screen; whatever sticks out to the
    ' right or below is cut off.

    If iRow < 0 Or iRow >= SCREEN_ROWS Or iCol < 0 Or iCol >= SCREEN_COLS Then
        MsgBox "Position (" & iRow & ", " & iCol & ") is off screen." & vbLf & _
               "Row must be 0.." & SCREEN_ROWS - 1 & ", column 0.." & SCREEN_COLS - 1 & ".", _
               vbExclamation, "PlaceShiftedSprite"
        Exit Sub
    End If

    ' Split the pixel column into byte column and shift, the way
    ' GET_BYTE_AND_SHIFT_FOR_COL does in the game (minus the $FF boundary).
    Dim lByteCol As Long: lByteCol = iCol \ PIXELS_PER_BYTE
    Dim lShift As Long: lShift = iCol Mod PIXELS_PER_BYTE

    ' Setting Shift makes 'Sprite Shifter' recompute BlockData.
    NamedRange("Shift").value = lShift
    If Application.Calculation = xlCalculationManual Then Application.Calculate

    ' Read BlockData once (11 x 3 hex strings, 1-based array) before writing.
    Dim vBlock As Variant: vBlock = NamedRange("BlockData").value

    Dim wsMem As Worksheet: Set wsMem = ThisWorkbook.Worksheets(MEMORY_SHEET)
    Dim rngAddr As Range: Set rngAddr = wsMem.Range(MEMORY_ADDR_RANGE)

    ' Each write would otherwise recalculate the whole screen, 33 times over.
    Dim lCalcMode As Long: lCalcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Cleanup

    Dim lR As Long, lB As Long
    Dim lScreenRow As Long, lMemRow As Long
    Dim vMatch As Variant

    For lR = 0 To SPRITE_ROWS - 1
        lScreenRow = iRow + lR
        If lScreenRow >= SCREEN_ROWS Then Exit For          ' cut off at the bottom

        vMatch = Application.match(Hex$(LineAddress(lScreenRow)), rngAddr, 0)
        If IsError(vMatch) Then
            MsgBox "Address of screen row " & lScreenRow & " not found in '" & _
                   MEMORY_SHEET & "'!" & MEMORY_ADDR_RANGE & ".", vbCritical, "PlaceShiftedSprite"
            GoTo Cleanup
        End If
        lMemRow = rngAddr.Row + vMatch - 1

        For lB = 0 To SPRITE_BYTES - 1
            If lByteCol + lB >= BYTES_PER_LINE Then Exit For  ' cut off at the right
            ' Written as text, the same format BlockData uses.
            wsMem.Cells(lMemRow, MEMORY_FIRST_DATA_COL + lByteCol + lB).value = _
                "'" & CStr(vBlock(lR + 1, lB + 1))
        Next lB
    Next lR

Cleanup:
    Application.Calculation = lCalcMode
    Application.ScreenUpdating = True
    If err.Number <> 0 Then
        MsgBox "Error " & err.Number & ": " & err.Description, vbCritical, "PlaceShiftedSprite"
    End If
End Sub


Public Function LineAddress(ByVal lScreenRow As Long) As Long
    ' Screen row 0..191 -> start address of that line on HGR page 1.
    ' Same lookup as 'Hires Pixels'!C: ROW_TO_OFFSET_HI OR $20, then ROW_TO_OFFSET_LO.
    ' The tables are 12 x 16, read row by row.
    Dim lTableRow As Long: lTableRow = lScreenRow \ OFFSET_TABLE_COLS + 1
    Dim lTableCol As Long: lTableCol = lScreenRow Mod OFFSET_TABLE_COLS + 1

    Dim lHi As Long
    lHi = HexToByte(NamedRange("RowToOffsetHiTable").Cells(lTableRow, lTableCol).value) Or HGR1_PAGE
    Dim lLo As Long
    lLo = HexToByte(NamedRange("RowToOffsetLoTable").Cells(lTableRow, lTableCol).value)

    LineAddress = lHi * 256 + lLo
End Function


Private Function NamedRange(ByVal sName As String) As Range
    ' Workbook-level name -> range, independent of the active sheet.
    Set NamedRange = ThisWorkbook.Names(sName).RefersToRange
End Function
