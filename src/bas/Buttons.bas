Attribute VB_Name = "Buttons"
' Copyright (c) 2026 Frank Schuhardt
' SPDX-License-Identifier: MIT

Option Explicit

'*************************************
' standard buttons
'*************************************

Sub Button_GenerateWorksheetsMatrix()
    GenerateWorksheetsMatrix
    UpdateDataFlowBoxes
End Sub

Sub Button_SetVersion()
    SetVersion
End Sub

Sub Button_UpdateDataFlowBoxes()
    UpdateDataFlowBoxes
End Sub


'*************************************
' local
'*************************************

Public Sub Button_UpdateViewer()
    UpdateViewer ActiveSheet
End Sub

Public Sub Button_LoadFromBytes()
    LoadFromBytes ActiveSheet
End Sub

Public Sub Button_LoadSpriteFromTable()
    LoadSpriteFromTable ActiveSheet
End Sub

Public Sub Button_HideHighBitColumns()
    Dim bHidden As Boolean: bHidden = Columns("I").Hidden
    Columns("I").Hidden = Not bHidden
    Columns("Q").Hidden = Not bHidden
End Sub

Public Sub Button_PlaceShiftedSprite()
    ' Top-left cell of the selection on 'Hires Screen' = sprite's top-left pixel.
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rngScreen As Range: Set rngScreen = ThisWorkbook.Names("HiresScreen").RefersToRange
    Dim rngCell As Range: Set rngCell = Selection.Cells(1)
    If rngCell.Worksheet.Name <> rngScreen.Worksheet.Name Then Exit Sub
    If Intersect(rngCell, rngScreen) Is Nothing Then
        MsgBox "Select a pixel cell inside the screen area.", vbExclamation, "Place shifted sprite"
        Exit Sub
    End If
    PlaceShiftedSprite rngCell.Row - rngScreen.Row, rngCell.Column - rngScreen.Column
End Sub



