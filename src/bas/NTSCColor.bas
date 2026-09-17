Attribute VB_Name = "NTSCColor"
Option Explicit

' NTSCColor -- NTSC artifact color rules, Chapter 3 page 7.
' Decision table: a 1-pixel with a 1-neighbor is white; a 1-pixel with
' no 1-neighbor is colored; a 0-pixel between two 1-neighbors is
' colored; every other 0-pixel is black.
'
' Color depends on the ABSOLUTE screen column (mod 2), not on a pixel's
' position within a sprite row -- the NTSC color-subcarrier phase is
' fixed relative to the physical screen, not to wherever a sprite
' happens to be drawn. This is why RowColors takes iBaseCol: the same
' bit pattern renders different colors depending on where on screen
' it lands, exactly as Lode Runner's own pixel-shift/column machinery
' (COMPUTE_SHIFTED_SPRITE etc.) has to account for.

Private Const COLOR_BLACK As Long = &H0        ' RGB(0,0,0)
Private Const COLOR_WHITE As Long = &HFFFFFF   ' RGB(255,255,255)

Private Const FIRST_COL As Long = 0
Private Const LAST_COL As Long = 13


Public Function ColoredPixel(ByVal iScreenCol As Long, ByVal iHighBit As Long) As Long
    ' Absolute screen column parity + high bit -> the one color pair
    ' selectable for that byte.
    Dim iParity As Long
    iParity = iScreenCol Mod 2
    If iHighBit = 0 Then
        If iParity = 0 Then
            ColoredPixel = RGB(255, 0, 255)     ' Violet
        Else
            ColoredPixel = RGB(0, 255, 0)       ' Green
        End If
    Else
        If iParity = 0 Then
            ColoredPixel = RGB(0, 192, 255)     ' Blue
        Else
            ColoredPixel = RGB(255, 192, 0)     ' Orange
        End If
    End If
End Function


Public Function PixelColor(ByVal iV As Long, ByVal iL As Long, ByVal iR As Long, _
                            ByVal iScreenCol As Long, ByVal iH As Long) As Long
    ' Decision table for one pixel. iL/iR are the true neighboring
    ' screen pixels (0 if genuinely off the left/right edge of HGR).
    If iV = 1 Then
        If iL = 1 Or iR = 1 Then
            PixelColor = COLOR_WHITE
        Else
            PixelColor = ColoredPixel(iScreenCol, iH)
        End If
    Else
        If iL = 1 And iR = 1 Then
            ' A colored 0 takes its hue from the flanking 1-bit
            ' responsible for it (its left neighbor), not from its own
            ' column. L and R always share one column parity (they are
            ' two columns apart); the 0 between them is the opposite
            ' parity. Coloring it from its own column fights that phase
            ' instead of continuing it -- and a repeating 1010101 byte
            ' would render as alternating stripes instead of the solid
            ' fill it's actually used for. Confirmed against the
            ' chapter's own worked sprite: every 0x55 row renders solid
            ' blue on page 8, not blue/orange dither.
            PixelColor = ColoredPixel(iScreenCol - 1, iH)
        Else
            PixelColor = COLOR_BLACK
        End If
    End If
End Function


Public Function RowColors(aiRow() As Long, _
                          ByVal iHB0 As Long, _
                          ByVal iHB1 As Long, _
                          ByVal iBaseCol As Long, _
                          Optional ByVal iLeftNeighbor As Long = 0, _
                          Optional ByVal iRightNeighbor As Long = 0) As Variant
    ' One sprite row: 14 pixels -> 14 RGB values.
    ' aiRow must be a 0-based array with 14 elements.
    '
    ' iBaseCol is the absolute Apple II hires screen column corresponding
    ' to aiRow(0), in the range 0-279. Deliverable 1's isolated-sprite
    ' viewer always passes 0, matching the chapter's own worked examples
    ' ("we place the sprite starting at column 0").
    '
    ' iLeftNeighbor is the screen pixel immediately left of aiRow(0).
    ' iRightNeighbor is the screen pixel immediately right of aiRow(13).
    ' Both default to 0 (true edge-of-screen / isolated-sprite behavior).
    '
    ' Columns 0-6 use iHB0, columns 7-13 use iHB1.

    Dim alColors(0 To 13) As Long
    Dim iCol As Long
    Dim iV As Long, iL As Long, iR As Long
    Dim iH As Long
    Dim iScreenCol As Long

    For iCol = 0 To 13
        iV = aiRow(iCol)

        If iCol > FIRST_COL Then
            iL = aiRow(iCol - 1)
        Else
            iL = iLeftNeighbor
        End If

        If iCol < LAST_COL Then
            iR = aiRow(iCol + 1)
        Else
            iR = iRightNeighbor
        End If

        If iCol < 7 Then
            iH = iHB0
        Else
            iH = iHB1
        End If

        iScreenCol = iBaseCol + iCol
        alColors(iCol) = PixelColor(iV, iL, iR, iScreenCol, iH)
    Next iCol

    RowColors = alColors
End Function
