Attribute VB_Name = "SpriteEditor"
Option Explicit

' SpriteEditor -- wires the Editor grid, Util, and NTSCColor together for
' the Sprite Editor/Viewer sheet.
'
' Two buttons, each a pure, idempotent function of its own inputs: every
' run fully overwrites every cell it owns, never accumulates, so running
' either button twice (or alternating them) on unchanged data is a no-op.
'
'   UpdateViewer  : pixels + HB (editor)  -> Hex0/Hex1, viewer colors
'   LoadFromBytes : Hex0/Hex1 (typed in)  -> pixels + HB (editor), viewer colors
'
' Bits0/Bits1 are NOT written here -- they are worksheet formulas
' (=HexToBits(...)) that follow Hex0/Hex1 automatically.

Private Const SHEET_NAME As String = "Sprite Editor-Viewer"
Private Const ROW_FIRST As Long = 4          ' editor/hex row for sprite row 0
Private Const VIEWER_ROW_FIRST As Long = 17  ' viewer row for sprite row 0
Private Const COL_HB0 As Long = 9            ' I
Private Const COL_HB1 As Long = 17           ' Q
Private Const COL_HEX0 As Long = 18          ' R
Private Const COL_HEX1 As Long = 19          ' S

Private Function PixelCol(ByVal iCol As Long) As Long
    ' Pixel columns 0-6 -> B-H (2-8); pixel columns 7-13 -> J-P (10-16).
    ' Column I (9) is the visual gap between the two bytes -- never a
    ' pixel column, so the +3 jump skips it.
    If iCol <= 6 Then
        PixelCol = iCol + 2
    Else
        PixelCol = iCol + 3
    End If
End Function

Public Function PixelsToByteValue(aiPixels() As Long, ByVal iRow As Long, _
                                   ByVal iByteIdx As Long, ByVal iHB As Long) As Long
    ' iByteIdx 0 -> pixels 0-6, iByteIdx 1 -> pixels 7-13.
    ' Bit 0 = leftmost pixel of the byte (LSB-first display, per the
    ' design doc's pixel<->byte convention). Bit 7 = high bit.
    Dim iBase As Long, iBit As Long, lVal As Long
    iBase = iByteIdx * 7
    lVal = 0
    For iBit = 0 To 6
        If aiPixels(iRow, iBase + iBit) = 1 Then
            lVal = SetBit(lVal, iBit, 1)
        End If
    Next iBit
    If iHB <> 0 Then lVal = SetBit(lVal, 7, 1)
    PixelsToByteValue = lVal
End Function

Public Sub UpdateViewer(ws As Worksheet)
    Dim aiPixels(0 To 10, 0 To 13) As Long
    Dim aiHB(0 To 10, 0 To 1) As Long
    Dim iRow As Long, iCol As Long

    ' 1. Read the editor grid (pixels + high bits) into arrays.
    For iRow = 0 To 10
        For iCol = 0 To 13
            aiPixels(iRow, iCol) = ws.Cells(ROW_FIRST + iRow, PixelCol(iCol)).Value
        Next iCol
        aiHB(iRow, 0) = ws.Cells(ROW_FIRST + iRow, COL_HB0).Value
        aiHB(iRow, 1) = ws.Cells(ROW_FIRST + iRow, COL_HB1).Value
    Next iRow

    ' 2. Compute and write Hex0/Hex1 for every row (Bits0/Bits1 follow
    '    automatically via the =HexToBits(...) formulas already in T/U).
    Dim lByte0 As Long, lByte1 As Long
    For iRow = 0 To 10
        lByte0 = PixelsToByteValue(aiPixels, iRow, 0, aiHB(iRow, 0))
        lByte1 = PixelsToByteValue(aiPixels, iRow, 1, aiHB(iRow, 1))
        ws.Cells(ROW_FIRST + iRow, COL_HEX0).Value = ByteToHex(lByte0)
        ws.Cells(ROW_FIRST + iRow, COL_HEX1).Value = ByteToHex(lByte1)
    Next iRow

    ' 3. Paint the viewer from the pixels just read.
    PaintViewer ws, aiPixels, aiHB
End Sub

Public Sub LoadFromBytes(ws As Worksheet)
    Dim aiPixels(0 To 10, 0 To 13) As Long
    Dim aiHB(0 To 10, 0 To 1) As Long
    Dim iRow As Long, iCol As Long
    Dim lByte0 As Long, lByte1 As Long

    ' 1. Read Hex0/Hex1 and decode into pixels + high bits.
    For iRow = 0 To 10
        lByte0 = HexToByte(ws.Cells(ROW_FIRST + iRow, COL_HEX0).Value)
        lByte1 = HexToByte(ws.Cells(ROW_FIRST + iRow, COL_HEX1).Value)

        aiHB(iRow, 0) = GetBit(lByte0, 7)
        aiHB(iRow, 1) = GetBit(lByte1, 7)

        For iCol = 0 To 6
            aiPixels(iRow, iCol) = GetBit(lByte0, iCol)
        Next iCol
        For iCol = 7 To 13
            aiPixels(iRow, iCol) = GetBit(lByte1, iCol - 7)
        Next iCol
    Next iRow

    ' 2. Write pixels + HB back into the editor grid -- every cell,
    '    every run, so this is idempotent regardless of prior state.
    For iRow = 0 To 10
        For iCol = 0 To 13
            ws.Cells(ROW_FIRST + iRow, PixelCol(iCol)).Value = aiPixels(iRow, iCol)
        Next iCol
        ws.Cells(ROW_FIRST + iRow, COL_HB0).Value = aiHB(iRow, 0)
        ws.Cells(ROW_FIRST + iRow, COL_HB1).Value = aiHB(iRow, 1)
    Next iRow

    ' 3. Paint the viewer from the pixels just decoded.
    PaintViewer ws, aiPixels, aiHB
End Sub

Public Sub LoadSpriteFromTable(ws As Worksheet)
    ' Pulls the 11 rows of the currently selected sprite (Sprite Loader!C4)
    ' out of SPRITE_DATA, using the addresses already computed in
    ' AddrByte0/AddrByte1 on "Sprite Loader". The high bit is forced on
    ' for every byte -- the stored table is essentially always high-bit-
    ' clear (see TODO.md), and the game itself sets the high bit via the
    ' shift/pattern tables at draw time, not per stored byte.
    '
    ' HexByte0/HexByte1 are local ranges (one per Sprite Editor-Viewer
    ' sheet), so the unqualified Range(...) calls below resolve to
    ' whichever such sheet is currently active -- same as typing into
    ' those cells by hand would.
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("SPRITE_DATA")

    Dim lBaseAddr As Long
    lBaseAddr = HexToByte(Range("SpriteBaseAddress").Value)

    Dim iRow As Long
    Dim lAddr As Long, lIdx As Long
    Dim iDataRow As Long, iDataCol As Long
    Dim lByte As Long

    For iRow = 1 To 11
        ' -- byte 1 --
        lAddr = HexToByte(Range("AddrByte0").Cells(iRow).Value)
        lIdx = lAddr - lBaseAddr
        iDataRow = 3 + lIdx \ 16
        iDataCol = 3 + (lIdx Mod 16)
        lByte = HexToByte(wsData.Cells(iDataRow, iDataCol).Value)
        lByte = SetBit(lByte, 7, 1)
        ws.Range("HexByte0").Cells(iRow).Value = ByteToHex(lByte)

        ' -- byte 2 --
        lAddr = HexToByte(Range("AddrByte1").Cells(iRow).Value)
        lIdx = lAddr - lBaseAddr
        iDataRow = 3 + lIdx \ 16
        iDataCol = 3 + (lIdx Mod 16)
        lByte = HexToByte(wsData.Cells(iDataRow, iDataCol).Value)
        lByte = SetBit(lByte, 7, 1)
        ws.Range("HexByte1").Cells(iRow).Value = ByteToHex(lByte)
    Next iRow

    ' Decompose into pixels + HB and paint the viewer, exactly once,
    ' after all 11 rows have been written.
    LoadFromBytes ws
End Sub

Private Sub PaintViewer(ByVal ws As Worksheet, aiPixels() As Long, aiHB() As Long)
    ' Shared by both buttons: given pixels + HB already in memory, repaint
    ' every viewer cell. iBaseCol is always 0 here -- Deliverable 1 shows
    ' the sprite in isolation, matching the chapter's own worked examples.
    Dim iRow As Long, iCol As Long
    Dim aiRow(0 To 13) As Long
    Dim vColors As Variant

    For iRow = 0 To 10
        For iCol = 0 To 13
            aiRow(iCol) = aiPixels(iRow, iCol)
        Next iCol

        vColors = RowColors(aiRow, aiHB(iRow, 0), aiHB(iRow, 1), 0)

        For iCol = 0 To 13
            ws.Cells(VIEWER_ROW_FIRST + iRow, PixelCol(iCol)).Interior.Color = vColors(iCol)
        Next iCol
    Next iRow
End Sub
