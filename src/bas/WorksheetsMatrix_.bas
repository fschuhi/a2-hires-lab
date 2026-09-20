Attribute VB_Name = "WorksheetsMatrix_"
Option Explicit

Const WS_Worksheets = "Worksheets matrix"
Const R_DataFlowBox = "DataFlowBox"
Const R_SkipWorksheets = "SkipWorksheets"
Const R_DotFile = "DotFile"

Const QUOTE = """"

'*************************************
' helpers
'*************************************

Function GetWorksheetsMatrixUpperLeft()
    Set GetWorksheetsMatrixUpperLeft = Worksheets(WS_Worksheets).Range("B3")
End Function

Function SkipWorksheets() As Dictionary
    Dim dictSkipped As New Dictionary
    Dim rSkip As Range
    For Each rSkip In Range(R_SkipWorksheets)
        If Not IsEmpty(rSkip) Then dictSkipped.Add rSkip.Value2, rSkip.Value2
    Next
    Set SkipWorksheets = dictSkipped
End Function


'*************************************
' generate matrix
'*************************************

Sub GenerateWorksheetsMatrix()
    Dim wsWorksheets As Worksheet: Set wsWorksheets = Worksheets(WS_Worksheets)
    Dim dictWorksheets As New Dictionary
    
    Dim rUpperLeft As Range: Set rUpperLeft = GetWorksheetsMatrixUpperLeft()
    ClearWorksheetsMatrix rUpperLeft
    
    Dim ix As Integer
    For ix = 1 To Worksheets.Count
        dictWorksheets.Add Worksheets(ix).Name, ix
    Next
    
    
    Dim ws As Worksheet
    Dim cRefs As Collection
    
    ' fill the matrix
    For ix = 1 To Worksheets.Count
        Set ws = Worksheets(ix)
            
        Set cRefs = SafeCollectSheetReferences(ws, dictWorksheets)
        If Not cRefs Is Nothing Then
            ShowStatus FS("# (#)", ws.Name, cRefs.Count)
            Dim vReferencedSheet
            For Each vReferencedSheet In cRefs
                Dim wsReferenced As Worksheet: Set wsReferenced = Worksheets(vReferencedSheet)
                Dim iReferencedWorksheetIndex As Integer: iReferencedWorksheetIndex = dictWorksheets(vReferencedSheet)
                SetReference rUpperLeft, ws, ix, wsReferenced, iReferencedWorksheetIndex
            Next
        End If
        
    Next
    
    ClearStatusBar
    
End Sub


Sub ClearWorksheetsMatrix(rUpperLeft As Range)
    Dim wsWorksheets As Worksheet: Set wsWorksheets = rUpperLeft.Worksheet
    
    Dim rLowerRight As Range: Set rLowerRight = Intersect(rUpperLeft.End(xlDown).End(xlDown).EntireRow, rUpperLeft.End(xlToRight).End(xlToRight).EntireColumn)
    Dim rLinks As Range: Set rLinks = wsWorksheets.Range(rUpperLeft, rLowerRight)
    
    rLinks.ClearContents
    
    On Error Resume Next
    rLinks.EntireColumn.Ungroup
    rLinks.EntireRow.Ungroup
    On Error GoTo 0
    
    ClearInteriorColor rLinks
    
    SetSilentApplicationState
    
    Dim ix As Integer
    Dim ws As Worksheet
    
    ShowStatus "building matrix..."
    
    Dim ixLastSection As Integer
    
    ' pass 1: build the matrix and the dict
    For ix = 1 To Worksheets.Count
        Set ws = Worksheets(ix)
        Dim rRow As Range: Set rRow = rUpperLeft.Offset(ix)
        Dim rCol As Range: Set rCol = rUpperLeft.Offset(0, ix)
        
        rRow = ws.Name
        rCol = ws.Name
        
        Dim rIntersect As Range: Set rIntersect = Intersect(rRow.EntireRow, rCol.EntireColumn)
        
        If IsSectionSheet(ws, False) Then
            If ixLastSection = 0 Then
                ' don't group anything
            Else
                wsWorksheets.Range(rUpperLeft.Offset(ixLastSection + 1), rRow.Offset(-1)).EntireRow.Group
                wsWorksheets.Range(rUpperLeft.Offset(0, ixLastSection + 1), rCol.Offset(0, -1)).EntireColumn.Group
            End If
            ixLastSection = ix
        End If
        
        ' set color of row and col header, and the intersection
        Dim col As Long: col = ws.Tab.Color
        If col = 0 Then
            rRow.Font.Color = BLACK
            rCol.Font.Color = BLACK
            ClearInteriorColor rRow
            ClearInteriorColor rCol
            rIntersect.Interior.Color = LIGHT_GRAY
        Else
            rRow.Font.Color = TextColorToUse(col)
            rCol.Font.Color = TextColorToUse(col)
            rRow.Interior.Color = col
            rCol.Interior.Color = col
            rIntersect.Interior.Color = col
        End If
        
    Next
    
    RevertApplicationState
End Sub


Function TextColorToUse(BackColor As Long) As Long
' https://www.mrexcel.com/board/threads/vba-read-only-tab-font-color.1141983/

'  This function returns the color to use for
'  text to make it readable on a dark background
'  Code by of Rick Rothstein
  Dim Luminance As Long
  Luminance = 77 * (BackColor Mod &H100) + _
              151 * ((BackColor \ &H100) Mod &H100) + _
              28 * ((BackColor \ &H10000) Mod &H100)
  '  Default value of TextColorToUse is 0-Black, set
  '  it to White if the Luminance is less than 32640
  If Luminance < 32640 Then TextColorToUse = vbWhite
End Function

'Sub dhancey()
'With Range("A1:C20")
'   .Interior.Color = .Parent.Tab.Color
'   .Font.Color = TextColorToUse(.Parent.Tab.Color)
'End With
'End Sub


Function SafeCollectSheetReferences(ws As Worksheet, dictSheets As Dictionary) As Collection
    Dim dictRefs As Dictionary
    On Error Resume Next
    'Set cRefs = CollectSheetReferences(ws.Cells.SpecialCells(xlCellTypeFormulas))
    Set dictRefs = CollectSheetReferencesDictionary_(ws.Cells, dictSheets)
    On Error GoTo 0
    
    If Not dictRefs Is Nothing Then
        
        Dim cSafeRefs As New Collection
        Dim wsReferenced As Worksheet
        Dim vSheetName
        For Each vSheetName In dictRefs.Keys
            Dim sSafeSheetName As String
            sSafeSheetName = Replace(vSheetName, "'", "")
            
            On Error Resume Next
            Set wsReferenced = Nothing: Set wsReferenced = Worksheets(sSafeSheetName)
            On Error GoTo 0
            If Not wsReferenced Is Nothing Then cSafeRefs.Add sSafeSheetName
        Next
        
        Set SafeCollectSheetReferences = cSafeRefs
    End If
End Function


Function CollectSheetReferencesDictionary_(rToSearchIn As Range, dictSheets As Dictionary) As Dictionary
    Dim rFormulas As Range
    On Error Resume Next
    Set rFormulas = rToSearchIn.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If rFormulas Is Nothing Then Exit Function
    
    Dim dRefs As New Dictionary
    Dim rFormula As Range
    For Each rFormula In rFormulas
        Dim sFormula As String: sFormula = rFormula.formula
        
        ' find ref
        Do
        
            Dim sRef As String
            
            sRef = RegMatch(sFormula, "('[^']+')!", 1)
        
            ' if not found a '...' reference then try to find a ref w/o the '
            If sRef = "" Then sRef = RegMatch(sFormula, "([^(=!]+)!", 1)
            
            ' 02.05.23 das funktioniert nicht, Excel bleibt stehen
            If sRef = "" Then
                sRef = RegMatch(sFormula, "([A-Za-z0-9_]+)\[[A-Za-z0-9() _-]+\]", 1)
                If sRef <> "" And Not dictSheets.Exists(sRef) Then Stop
            End If
            
            ' add the ref if not already found earlier
            If sRef <> "" Then
            
                ' this differs from above CollectSheetReferences()
                Dim cRefs As Collection
                If Not dRefs.Exists(sRef) Then
                    Set cRefs = New Collection
                    dRefs.Add sRef, cRefs
                Else
                    Set cRefs = dRefs(sRef)
                End If
                cRefs.Add rFormula
                                
                'sFormula = Replace(sFormula, sRef & "!", "")
                sFormula = Replace(sFormula, LastMatch0, "")
            End If
        Loop Until sRef = ""
    Next

    
    Set CollectSheetReferencesDictionary_ = dRefs
End Function


Sub SetReference(rUpperLeft As Range, wsThis As Worksheet, iThisIndex As Integer, wsReferenced As Worksheet, iReferencedIndex As Integer)
    Dim rReferenceIn As Range: Set rReferenceIn = rUpperLeft.Offset(iThisIndex, iReferencedIndex)
    Dim rReferenceOut As Range: Set rReferenceOut = rUpperLeft.Offset(iReferencedIndex, iThisIndex)
    
    If IsEmpty(rReferenceIn) Then
        rReferenceIn.formula = FS("=LinkSheets(""<="", '#'!A1, '#'!A1)", wsThis.Name, wsReferenced.Name)
        'rReferenceIn = "<="
    Else
        rReferenceIn.formula = FS("=LinkSheets(""<=>"", '#'!A1, '#'!A1)", wsThis.Name, wsReferenced.Name)
    End If
    If IsEmpty(rReferenceOut) Then
        rReferenceOut.formula = FS("=LinkSheets(""=>"", '#'!A1, '#'!A1)", wsThis.Name, wsReferenced.Name)
    Else
        rReferenceOut.formula = FS("=LinkSheets(""<=>"", '#'!A1, '#'!A1)", wsThis.Name, wsReferenced.Name)
    End If
End Sub


Function LinkSheets(s As String, r1 As Range, r2 As Range)
    LinkSheets = s
End Function


'*************************************
' update dataflow boxes
'*************************************

Sub UpdateDataFlowBoxes()
    Dim dictSkipped As Dictionary: Set dictSkipped = SkipWorksheets()

    Dim wsWorksheets As Worksheet: Set wsWorksheets = Worksheets(WS_Worksheets)
    Dim rUpperLeft As Range: Set rUpperLeft = GetWorksheetsMatrixUpperLeft()
    Dim rWorksheets As Range: Set rWorksheets = wsWorksheets.Range(rUpperLeft.Offset(1), rUpperLeft.Offset(1).End(xlDown))
    Dim ix
    Dim sh
    Dim cLines As Collection
    For ix = 1 To rWorksheets.Count
        Dim sWorksheetName As String: sWorksheetName = rWorksheets.Cells(ix)
        Dim ws As Worksheet: Set ws = Nothing
        On Error Resume Next
        Set ws = Worksheets(sWorksheetName)
        On Error GoTo 0
        
        If Not ws Is Nothing Then
        
            Set sh = Nothing
            On Error Resume Next
            Set sh = ws.Shapes.Range(R_DataFlowBox)
            On Error GoTo 0
            If Not sh Is Nothing Then
            
                Set cLines = New Collection
                Dim ixRef
                For ixRef = 1 To rWorksheets.Count
                    Dim sReferencedSheet As String: sReferencedSheet = rWorksheets.Cells(ixRef)
                    If dictSkipped.Exists(sReferencedSheet) Then
                        ' we don't want to show this sheet in the box
                    Else
                        Dim rRef As Range: Set rRef = rUpperLeft.Offset(ix, ixRef)
                        If (Not IsEmpty(rRef)) And (Not IsError(rRef)) Then
                        'If Not IsEmpty(rRef) Then
                            cLines.Add rRef & " " & sReferencedSheet
                        End If
                    End If
                Next
            
                Dim sLines As String: sLines = CollectionToString(cLines, Chr(13) & Chr(10))
                If sLines = "" Then sLines = "(no references)"
                sh.TextFrame2.textRange.Characters.Text = sLines
            End If
        End If
    Next
End Sub


Sub UpdateDataFlowBox(ws As Worksheet, cLines As Collection)
    Dim sLines As String: sLines = CollectionToString(cLines, Chr(13) & Chr(10))
    ws.Shapes.Range(R_DataFlowBox).TextFrame2.textRange.Characters.Text = sLines
End Sub


'*************************************
' Graphviz dot
'*************************************

Sub GenerateDot()
    Dim wsWorksheets As Worksheet: Set wsWorksheets = Worksheets(WS_Worksheets)
    Dim rUpperLeft As Range: Set rUpperLeft = GetWorksheetsMatrixUpperLeft()
    Dim rRows As Range: Set rRows = wsWorksheets.Range(rUpperLeft.Offset(1, 0), rUpperLeft.Offset(1, 0).End(xlDown))
    Dim rCols As Range: Set rCols = wsWorksheets.Range(rUpperLeft.Offset(0, 1), rUpperLeft.Offset(0, 1).End(xlToRight))
    Dim dictSkipped As Dictionary: Set dictSkipped = SkipWorksheets()
    
    Dim cLines As New Collection
    cLines.Add "digraph G {"
    cLines.Add FS("node [shape=box, color=#darkgray#, fontname=#Arial#, fontsize=8];", QUOTE, QUOTE, QUOTE, QUOTE)
    cLines.Add FS("edge [arrowhead=vee, color=#darkgray#, arrowsize=0.5];", QUOTE, QUOTE)
        
    Dim rRow As Range
    Dim rCol As Range
    Dim ixRow
    For ixRow = 1 To rRows.Count
        Set rRow = rRows.Cells(ixRow)
        Dim sRow As String: sRow = rRow
        If Not dictSkipped.Exists(sRow) Then
        
            Dim ixCol
            For ixCol = ixRow To rCols.Count
                Set rCol = rCols.Cells(ixCol)
                Dim sCol As String: sCol = rCol
                If Not dictSkipped.Exists(sCol) Then
                    Dim rRel As Range: Set rRel = Intersect(rRow.EntireRow, rCol.EntireColumn)
                    If Not IsEmpty(rRel) Then
                        Dim sDotRel As String: sDotRel = QUOTE & sRow & QUOTE & " -> " & QUOTE & sCol & QUOTE
                        Select Case rRel.Value2
                            Case "=>"
                                cLines.Add sDotRel & " [dir=forward];"
                            Case "<="
                                cLines.Add sDotRel & " [dir=back];"
                            Case "<=>"
                                cLines.Add sDotRel & " [dir=both];"
                        End Select
                    End If
                End If
            Next
        End If
    Next
    
    cLines.Add "}"
    SaveCollectionToTextFile cLines, Range(R_DotFile)
End Sub



