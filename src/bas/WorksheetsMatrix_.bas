Attribute VB_Name = "WorksheetsMatrix_"
' Copyright (c) 2009-2026 Frank Schuhardt
' SPDX-License-Identifier: MIT

Option Explicit

Const WS_Worksheets = "Worksheets matrix"
Const R_DataFlowBox = "DataFlowBox"
Const R_SkipWorksheets = "SkipWorksheets"
Const R_DotFile = "DotFile"

Const QUOTE = """"

' reference scanner, see InitReferenceScanner
Private Const MAX_NAME_DEPTH = 20

Private m_reString As RegExp            ' string literals, removed before scanning
Private m_reSheetRef As RegExp          ' Sheet! / 'Sheet'! / [Book]Sheet!
Private m_reToken As RegExp             ' identifiers that might be names or tables
Private m_reBrackets As RegExp          ' [...] contents: table columns, R1C1 offsets
Private m_dSheetsLC As Dictionary       ' LCase(sheet name) -> sheet name
Private m_dNameDefs As Dictionary       ' "name" or "sheet!name" (LCase) -> Array(RefersTo, context sheet)
Private m_dNameSheets As Dictionary     ' same keys -> Dictionary of sheet names (resolved, memoized)
Private m_dTableSheets As Dictionary    ' LCase(table name) -> sheet name

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
    
    InitReferenceScanner
    SetSilentApplicationState
    
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
    
    RevertApplicationState
    ClearReferenceScanner
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


'*************************************
' reference scanner
'*************************************

' Finds the sheets a formula depends on:
'   - explicit sheet references: Sheet!A1, 'My Sheet'!A1, 'Bob''s'!A1, Sheet1:Sheet3!A1
'     (surrounding whitespace and line breaks don't matter; external workbooks are ignored)
'   - defined names, global and sheet-local, including names defined via other names
'     or via formulas (e.g. =OFFSET(Data!$A$1, ...))
'   - Excel tables (ListObjects), e.g. Table1[Column] or ROWS(Table1)
' Not visible to any formula scan: INDIRECT with computed text, and data moved by VBA.

Private Sub InitReferenceScanner()
    Set m_reString = New RegExp
    m_reString.Global = True
    m_reString.PATTERN = """(?:[^""]|"""")*"""
    
    ' group 1: external workbook prefix (skipped), group 2: quoted name, group 3: unquoted name (maybe 3D)
    Const UNQUOTED = "[^\s'!(),=+\-*/&^<>;:{}""\[\]%#@$]+"
    Set m_reSheetRef = New RegExp
    m_reSheetRef.Global = True
    m_reSheetRef.PATTERN = "(\[[^\]]*\])?(?:'((?:[^']|'')+)'|(" & UNQUOTED & "(?::" & UNQUOTED & ")?))!"
    
    Set m_reBrackets = New RegExp
    m_reBrackets.Global = True
    m_reBrackets.PATTERN = "\[[^\[\]]*\]"
    
    Set m_reToken = New RegExp
    m_reToken.Global = True
    m_reToken.PATTERN = "[A-Za-z_\\\u00C0-\u024F][A-Za-z0-9_.\\?\u00C0-\u024F]*"
    
    Set m_dSheetsLC = New Dictionary
    Dim ws As Worksheet
    For Each ws In Worksheets
        m_dSheetsLC(LCase(ws.Name)) = ws.Name
    Next
    
    Set m_dTableSheets = New Dictionary
    Dim lo As ListObject
    For Each ws In Worksheets
        For Each lo In ws.ListObjects
            m_dTableSheets(LCase(lo.Name)) = ws.Name
        Next
    Next
    
    Set m_dNameDefs = New Dictionary
    Set m_dNameSheets = New Dictionary
    Dim nm As Name
    For Each nm In ActiveWorkbook.Names
        Dim sBare As String: sBare = Mid(nm.Name, InStrRev(nm.Name, "!") + 1)
        ' skip Excel's internal names (_xlnm._FilterDatabase, _xlpm.* LET parameters, ...)
        If LCase(Left(sBare, 3)) <> "_xl" Then
            Dim sRefersTo As String: sRefersTo = ""
            On Error Resume Next
            sRefersTo = nm.RefersTo
            On Error GoTo 0
            
            If TypeName(nm.Parent) = "Worksheet" Then
                m_dNameDefs(LCase(nm.Parent.Name & "!" & sBare)) = Array(sRefersTo, nm.Parent.Name)
            Else
                m_dNameDefs(LCase(sBare)) = Array(sRefersTo, "")
            End If
        End If
    Next
End Sub


Private Sub ClearReferenceScanner()
    Set m_dSheetsLC = Nothing
    Set m_dNameDefs = Nothing
    Set m_dNameSheets = Nothing
    Set m_dTableSheets = Nothing
End Sub


Private Sub AddSheetRef(ByVal sSheet As String, dOut As Dictionary)
    Dim sKey As String: sKey = LCase(Trim(sSheet))
    If m_dSheetsLC.Exists(sKey) Then dOut(m_dSheetsLC(sKey)) = True
End Sub


Private Function ResolveName(sKey As String, iDepth As Integer) As Dictionary
    ' sheets a defined name points to; memoized, and safe against names defined via each other
    If m_dNameSheets.Exists(sKey) Then
        Set ResolveName = m_dNameSheets(sKey)
        Exit Function
    End If
    
    Dim dSheets As Dictionary: Set dSheets = New Dictionary
    Set m_dNameSheets(sKey) = dSheets       ' register first, so cycles terminate
    
    Dim vDef: vDef = m_dNameDefs(sKey)
    CollectRefsFromFormula CStr(vDef(0)), CStr(vDef(1)), dSheets, iDepth + 1
    
    Set ResolveName = dSheets
End Function


Private Sub AddNameOrTableRef(sToken As String, sContextSheet As String, dOut As Dictionary, iDepth As Integer)
    ' a sheet-local name wins over a global name of the same name
    Dim sKey As String: sKey = LCase(sContextSheet & "!" & sToken)
    If Not m_dNameDefs.Exists(sKey) Then sKey = LCase(sToken)
    
    If m_dNameDefs.Exists(sKey) Then
        Dim vSheet
        For Each vSheet In ResolveName(sKey, iDepth).Keys
            dOut(vSheet) = True
        Next
    ElseIf m_dTableSheets.Exists(LCase(sToken)) Then
        dOut(m_dTableSheets(LCase(sToken))) = True
    End If
End Sub


Private Sub CollectRefsFromFormula(ByVal sFormula As String, ByVal sContextSheet As String, dOut As Dictionary, ByVal iDepth As Integer)
    ' adds every sheet sFormula depends on to dOut (as keys)
    If iDepth > MAX_NAME_DEPTH Then Exit Sub
    If sFormula = "" Then Exit Sub
    
    sFormula = m_reString.Replace(sFormula, """""")
    
    ' 1) explicit sheet references
    Dim m, vPart
    For Each m In m_reSheetRef.Execute(sFormula)
        If IsEmpty(m.SubMatches(0)) Or m.SubMatches(0) = "" Then     ' not an external [Book]Sheet! reference
            Dim sSheetPart As String
            If m.SubMatches(1) <> "" Then
                sSheetPart = Replace(m.SubMatches(1), "''", "'")
                If InStr(sSheetPart, "]") > 0 Then sSheetPart = ""    ' quoted external reference
            Else
                sSheetPart = m.SubMatches(2)
            End If
            
            ' 3D references (Sheet1:Sheet3!A1): record both ends
            If sSheetPart <> "" Then
                For Each vPart In Split(sSheetPart, ":")
                    AddSheetRef CStr(vPart), dOut
                Next
            End If
        End If
    Next
    
    ' 2) defined names and tables
    '    sheet prefixes are collapsed to "!", so the cell/name behind them is skipped below
    '    bracket contents (table columns like [@Amount], R1C1 offsets) are dropped, table names stay
    Dim sRest As String: sRest = m_reSheetRef.Replace(sFormula, "!")
    sRest = m_reBrackets.Replace(sRest, "")
    sRest = m_reBrackets.Replace(sRest, "")     ' second pass for nested [[#This Row],[Amount]]
    For Each m In m_reToken.Execute(sRest)
        Dim iStart As Long: iStart = m.FirstIndex + 1
        Dim sPrev As String: sPrev = ""
        If iStart > 1 Then sPrev = Mid(sRest, iStart - 1, 1)
        Dim sNext As String: sNext = Mid(sRest, iStart + m.Length, 1)
        
        ' not: Sheet!X (already covered), functions X(, absolute column $A$1
        If sPrev <> "!" And sNext <> "(" And sNext <> "!" And sNext <> "$" Then
            AddNameOrTableRef CStr(m.value), sContextSheet, dOut, iDepth
        End If
    Next
End Sub


Function SafeCollectSheetReferences(ws As Worksheet, dictSheets As Dictionary) As Collection
    ' names of the sheets ws depends on, without ws itself
    Dim dictRefs As Dictionary
    Set dictRefs = CollectSheetReferencesDictionary_(ws.Cells)
    If dictRefs Is Nothing Then Exit Function
    
    Dim cSafeRefs As New Collection
    Dim vSheetName
    For Each vSheetName In dictRefs.Keys
        If vSheetName <> ws.Name And dictSheets.Exists(vSheetName) Then cSafeRefs.Add vSheetName
    Next
    
    Set SafeCollectSheetReferences = cSafeRefs
End Function


Function CollectSheetReferencesDictionary_(rToSearchIn As Range) As Dictionary
    ' referenced sheet name -> Collection of the cells referencing it
    ' (one cell per distinct formula: filled-down formulas are identical in R1C1 and scanned once)
    If m_dSheetsLC Is Nothing Then InitReferenceScanner
    
    Dim rFormulas As Range
    On Error Resume Next
    Set rFormulas = rToSearchIn.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If rFormulas Is Nothing Then Exit Function
    
    Dim sContextSheet As String: sContextSheet = rToSearchIn.Worksheet.Name
    Dim dSeen As New Dictionary
    Dim dRefs As New Dictionary
    Dim rFormula As Range
    For Each rFormula In rFormulas
        Dim sFormula As String: sFormula = rFormula.FormulaR1C1
        If Not dSeen.Exists(sFormula) Then
            dSeen.Add sFormula, True
            
            Dim dCell As Dictionary: Set dCell = New Dictionary
            CollectRefsFromFormula sFormula, sContextSheet, dCell, 0
            
            Dim vSheet
            For Each vSheet In dCell.Keys
                If Not dRefs.Exists(vSheet) Then dRefs.Add vSheet, New Collection
                dRefs(vSheet).Add rFormula
            Next
        End If
    Next
    
    Set CollectSheetReferencesDictionary_ = dRefs
End Function


Sub SetReference(rUpperLeft As Range, wsThis As Worksheet, iThisIndex As Integer, wsReferenced As Worksheet, iReferencedIndex As Integer)
    Dim rReferenceIn As Range: Set rReferenceIn = rUpperLeft.Offset(iThisIndex, iReferencedIndex)
    Dim rReferenceOut As Range: Set rReferenceOut = rUpperLeft.Offset(iReferencedIndex, iThisIndex)
    
    If IsEmpty(rReferenceIn) Then
        rReferenceIn.Formula = FS("=LinkSheets(""<="", '#'!A1, '#'!A1)", wsThis.Name, wsReferenced.Name)
        'rReferenceIn = "<="
    Else
        rReferenceIn.Formula = FS("=LinkSheets(""<=>"", '#'!A1, '#'!A1)", wsThis.Name, wsReferenced.Name)
    End If
    If IsEmpty(rReferenceOut) Then
        rReferenceOut.Formula = FS("=LinkSheets(""=>"", '#'!A1, '#'!A1)", wsThis.Name, wsReferenced.Name)
    Else
        rReferenceOut.Formula = FS("=LinkSheets(""<=>"", '#'!A1, '#'!A1)", wsThis.Name, wsReferenced.Name)
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
