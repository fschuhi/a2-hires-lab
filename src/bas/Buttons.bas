Attribute VB_Name = "Buttons"
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


