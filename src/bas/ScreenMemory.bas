Attribute VB_Name = "ScreenMemory"
' Copyright (c) 2026 Frank Schuhardt
' SPDX-License-Identifier: MIT

Option Explicit

' ScreenMemory -- the HGR page 1 memory model on 'Hires Memory' and its
' coloured picture on 'Hires Screen'.
'
' 'Hires Memory' lists the 192 screen lines in address order (C3:C194),
' with the 40 bytes of each line in F:AS. It is the single source of truth.
' 'Hires Screen' has no formulas: PaintScreen colours its cells from memory.
' 'Hires Screen (debug)', 'Hires Pixels' and 'Hires HB' show the same bytes
' through live formulas, for inspection.
'
' Coordinates are 0-based, like the labels on 'Hires Screen' and the 6502:
' screen rows 0..191, pixel columns 0..279.
'
' PlaceShiftedSprite is a first step only: it ignores the $FF boundary
' problem the game has with 280 columns, does no masking and does not
' clear what was there before.

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

Private Const NTSC_BLACK As Long = &H0               ' what PixelColor returns for black
Private Const SCREEN_BLACK As Long = &HBFBFBF        ' RGB(191,191,191), shown instead of black


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

    ' Each write would otherwise recalculate the debug sheets, 33 times over.
    SetSilentApplicationState
    On Error GoTo Cleanup

    Dim lR As Long, lB As Long
    Dim lScreenRow As Long, lMemRow As Long

    For lR = 0 To SPRITE_ROWS - 1
        lScreenRow = iRow + lR
        If lScreenRow >= SCREEN_ROWS Then Exit For          ' cut off at the bottom

        lMemRow = MemoryRow(wsMem, lScreenRow)

        For lB = 0 To SPRITE_BYTES - 1
            If lByteCol + lB >= BYTES_PER_LINE Then Exit For  ' cut off at the right
            ' Written as text, the same format BlockData uses.
            wsMem.Cells(lMemRow, MEMORY_FIRST_DATA_COL + lByteCol + lB).value = _
                "'" & CStr(vBlock(lR + 1, lB + 1))
        Next lB
    Next lR

    ' Repaint what the sprite covered, plus one pixel left and right: those
    ' neighbours can change colour even though the sprite didn't write them.
    PaintScreen iRow, iRow + SPRITE_ROWS - 1, _
                lByteCol * PIXELS_PER_BYTE - 1, (lByteCol + SPRITE_BYTES) * PIXELS_PER_BYTE

Cleanup:
    RevertApplicationState
    If err.Number <> 0 Then
        MsgBox "Error " & err.Number & ": " & err.Description, vbCritical, "PlaceShiftedSprite"
    End If
End Sub


Public Sub PaintScreen(ByVal lRowFrom As Long, ByVal lRowTo As Long, _
                       ByVal lColFrom As Long, ByVal lColTo As Long)
    ' The Apple II video circuit -- Apple II, not Lode Runner.
    '
    ' On the real machine nothing "draws" the screen. The video hardware reads
    ' HGR page 1 ($2000-$3FFF) line by line, 60 times a second, and turns the
    ' bits into the TV signal. Each byte gives 7 pixels, bit 0 leftmost; bit 7
    ' is not a pixel but selects the colour pair for that byte. What colour a
    ' TV shows follows the NTSC rules in NTSCColor.PixelColor: the pixel's own
    ' bit, its left and right neighbours, the parity of its absolute screen
    ' column, and its own byte's high bit.
    '
    ' PaintScreen does the same for a rectangle of 'Hires Screen' (rows and
    ' pixel columns, inclusive, 0-based): it reads those lines from
    ' 'Hires Memory' and colours each cell. It knows nothing about sprites --
    ' any code that writes memory can call it for the area it changed.
    '
    ' Neighbours always come from memory, even at the edge of the rectangle,
    ' so painting a part gives the same colours as painting everything. The
    ' caller has to include one pixel on each side of what it wrote, because
    ' those neighbours may change colour. Off-screen neighbours count as 0.
    '
    ' Black is shown as SCREEN_BLACK (light grey): easier on the eyes, and
    ' white pixels still stand out.

    ' Clip to the screen.
    If lRowFrom < 0 Then lRowFrom = 0
    If lRowTo > SCREEN_ROWS - 1 Then lRowTo = SCREEN_ROWS - 1
    If lColFrom < 0 Then lColFrom = 0
    If lColTo > SCREEN_COLS - 1 Then lColTo = SCREEN_COLS - 1
    If lRowFrom > lRowTo Or lColFrom > lColTo Then Exit Sub

    Dim rngScreen As Range: Set rngScreen = NamedRange("HiresScreen")
    Dim wsMem As Worksheet: Set wsMem = ThisWorkbook.Worksheets(MEMORY_SHEET)

    SetSilentApplicationState
    On Error GoTo Cleanup

    ' Background in one call; below, only non-black runs are painted.
    rngScreen.Cells(lRowFrom + 1, lColFrom + 1) _
             .Resize(lRowTo - lRowFrom + 1, lColTo - lColFrom + 1).Interior.Color = SCREEN_BLACK

    Dim alPixel(0 To SCREEN_COLS - 1) As Long       ' 1 = lit, in screen order
    Dim alHigh(0 To BYTES_PER_LINE - 1) As Long     ' bit 7 of each byte
    Dim vLine As Variant
    Dim lY As Long, lB As Long, lK As Long, lX As Long
    Dim lByte As Long, lL As Long, lR As Long
    Dim lColor As Long, lRunColor As Long, lRunStart As Long

    For lY = lRowFrom To lRowTo
        ' Decode the whole line: 40 bytes -> 280 pixels + 40 high bits.
        vLine = wsMem.Cells(MemoryRow(wsMem, lY), MEMORY_FIRST_DATA_COL) _
                     .Resize(1, BYTES_PER_LINE).value
        For lB = 0 To BYTES_PER_LINE - 1
            lByte = CellByte(vLine(1, lB + 1))
            alHigh(lB) = GetBit(lByte, 7)
            For lK = 0 To PIXELS_PER_BYTE - 1
                alPixel(lB * PIXELS_PER_BYTE + lK) = GetBit(lByte, lK)
            Next lK
        Next lB

        ' Colour pixel by pixel, write runs of the same colour.
        For lX = lColFrom To lColTo
            If lX > 0 Then lL = alPixel(lX - 1) Else lL = 0
            If lX < SCREEN_COLS - 1 Then lR = alPixel(lX + 1) Else lR = 0
            lColor = PixelColor(alPixel(lX), lL, lR, lX, alHigh(lX \ PIXELS_PER_BYTE))
            If lColor = NTSC_BLACK Then lColor = SCREEN_BLACK

            If lX = lColFrom Then
                lRunStart = lX: lRunColor = lColor
            ElseIf lColor <> lRunColor Then
                FillRun rngScreen, lY, lRunStart, lX - 1, lRunColor
                lRunStart = lX: lRunColor = lColor
            End If
        Next lX
        FillRun rngScreen, lY, lRunStart, lColTo, lRunColor
    Next lY

Cleanup:
    RevertApplicationState
    If err.Number <> 0 Then
        MsgBox "Error " & err.Number & ": " & err.Description, vbCritical, "PaintScreen"
    End If
End Sub


Public Sub ClearScreen()
    ' Like HGR on the Apple II: zero page 1, and the screen goes black.
    ' Empty cells in 'Hires Memory' read as $00.
    SetSilentApplicationState
    NamedRange("HiresMemory").ClearContents
    NamedRange("HiresScreen").Interior.Color = SCREEN_BLACK
    RevertApplicationState
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


Private Function MemoryRow(ByVal wsMem As Worksheet, ByVal lScreenRow As Long) As Long
    ' Screen row -> sheet row of that line on 'Hires Memory' (found by address).
    Dim rngAddr As Range: Set rngAddr = wsMem.Range(MEMORY_ADDR_RANGE)
    Dim vMatch As Variant
    vMatch = Application.match(Hex$(LineAddress(lScreenRow)), rngAddr, 0)
    If IsError(vMatch) Then
        err.Raise vbObjectError + 514, , "Address of screen row " & lScreenRow & _
                  " not found in '" & MEMORY_SHEET & "'!" & MEMORY_ADDR_RANGE & "."
    End If
    MemoryRow = rngAddr.Row + vMatch - 1
End Function


Private Function CellByte(ByVal vCell As Variant) As Long
    ' One 'Hires Memory' cell -> byte. Empty = $00. A typed number such as 33
    ' is read as hex "33", the same way HEX2BIN reads it on the debug sheets.
    If IsEmpty(vCell) Then
        CellByte = 0
    ElseIf Len(Trim$(CStr(vCell))) = 0 Then
        CellByte = 0
    Else
        CellByte = HexToByte(CStr(vCell))
    End If
End Function


Private Sub FillRun(ByVal rngScreen As Range, ByVal lY As Long, _
                    ByVal lXFrom As Long, ByVal lXTo As Long, ByVal lColor As Long)
    ' One run of same-coloured pixels. Black runs are already painted.
    If lColor = SCREEN_BLACK Then Exit Sub
    rngScreen.Cells(lY + 1, lXFrom + 1).Resize(1, lXTo - lXFrom + 1).Interior.Color = lColor
End Sub


Private Function NamedRange(ByVal sName As String) As Range
    ' Workbook-level name -> range, independent of the active sheet.
    Set NamedRange = ThisWorkbook.Names(sName).RefersToRange
End Function
