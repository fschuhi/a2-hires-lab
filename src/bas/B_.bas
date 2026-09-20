Attribute VB_Name = "B_"
' Copyright (c) 2009-2026 Frank Schuhardt, frank.schuhardt@gmail.com

' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Hiermit wird unentgeltlich jeder Person, die eine Kopie der Software und der zugehörigen Dokumentationen (die "Software") erhält, die Erlaubnis erteilt,
' sie uneingeschränkt zu nutzen, inklusive und ohne Ausnahme mit dem Recht, sie zu verwenden, zu kopieren, zu verändern, zusammenzufügen, zu veröffentlichen,
' zu verbreiten, zu unterlizenzieren und/oder zu verkaufen, und Personen, denen diese Software überlassen wird, diese Rechte zu verschaffen,
' unter den folgenden Bedingungen:

' Der obige Urheberrechtsvermerk und dieser Erlaubnisvermerk sind in allen Kopien oder Teilkopien der Software beizulegen.

' DIE SOFTWARE WIRD OHNE JEDE AUSDRÜCKLICHE ODER IMPLIZIERTE GARANTIE BEREITGESTELLT, EINSCHLIESSLICH DER GARANTIE ZUR BENUTZUNG FÜR DEN VORGESEHENEN ODER EINEM
' BESTIMMTEN ZWECK SOWIE JEGLICHER RECHTSVERLETZUNG, JEDOCH NICHT DARAUF BESCHRÄNKT. IN KEINEM FALL SIND DIE AUTOREN ODER COPYRIGHTINHABER FÜR JEGLICHEN SCHADEN
' ODER SONSTIGE ANSPRÜCHE HAFTBAR ZU MACHEN, OB INFOLGE DER ERFÜLLUNG EINES VERTRAGES, EINES DELIKTES ODER ANDERS IM ZUSAMMENHANG MIT DER SOFTWARE ODER
' SONSTIGER VERWENDUNG DER SOFTWARE ENTSTANDEN.

' 20.09.26 A_ -> B_ as baseline for `a2-hires-lab`

Option Explicit


' ApplicationStates

' prevent Worksheet_Change
Public WorksheetChangeDisabled As Boolean

Private g_cApplicationStates As Collection


' Comments

Const WS_COMMENTS_DEFAULT = "> comments <"

Public g_iUserFormCommentTop As Integer
Public g_iUserFormCommentLeft As Integer


' Names

Public Const R_NoCommentsWorksheets = "NoCommentsWorksheets"


' RegExpr

' buffer regular expressions in a dictionary, to speed up mass RegMatch()
Public g_dRegExps As Dictionary

Private m_LastMatches As Object
Private m_LastTarget As Variant
Private m_LastPattern As String
Public LastMatch0 As String


' SheetJump

Public Const R_SectionMidpoint = "SectionMidpoint"
Public Const DEFAULT_SECTION_MIDPOINT = 30


' Versions

Const R_VersionBox = "VersionBox"



'>> Addresses

Public Function GetAddress(r As Range, Optional bWithSheet As Boolean) As String
    If Not bWithSheet Then
        GetAddress = r.Address
    Else
        Dim sAddress As String: sAddress = r.Address(External:=True)
        If r.Worksheet.Parent Is ActiveWorkbook Then
            Dim ix1 As Integer: ix1 = InStr(sAddress, "[")
            Dim ix2 As Integer: ix2 = InStr(sAddress, "]")
            sAddress = Left(sAddress, ix1 - 1) & Mid(sAddress, ix2 + 1)
        End If
        GetAddress = sAddress
    End If
End Function



'>> ApplicationStates

' That's what you usually call to isolate a sub or function from time-consuming recalcs and screen updates

Sub SetSilentApplicationState()
    SaveApplicationState
    If Application.Calculation <> xlCalculationManual Then Application.Calculation = xlCalculationManual
    If Application.ScreenUpdating = True Then Application.ScreenUpdating = False
    If WorksheetChangeDisabled = False Then WorksheetChangeDisabled = True
End Sub

' That's what to call before leaving the sub/function which called SetSilentApplicationState
' VBA doesn't have try...catch, so you will usually want to make sure that you have an OnError active.

Sub RevertApplicationState()
    If SafeGetApplicationStates.Count = 0 Then Exit Sub
    Dim v
    
    v = Pop(): If v <> WorksheetChangeDisabled Then WorksheetChangeDisabled = v
    v = Pop(): If v <> Application.Calculation Then Application.Calculation = v
    v = Pop(): If v <> Application.ScreenUpdating Then Application.ScreenUpdating = v
    Application.StatusBar = PopStatusBar()
End Sub


'*************************************
' SafeGetApplicationStates
'*************************************

Private Sub ForceInitApplicationStates()
    Set g_cApplicationStates = New Collection
End Sub

Private Function SafeGetApplicationStates() As Collection
    If g_cApplicationStates Is Nothing Then ForceInitApplicationStates
    Set SafeGetApplicationStates = g_cApplicationStates
End Function


'*************************************
' push / pop
'*************************************

Private Sub SaveApplicationState()
    ' pushes certain states on the "stack"
    
    Push Application.StatusBar
    Push Application.ScreenUpdating
    Push Application.Calculation
    Push WorksheetChangeDisabled
End Sub

Private Sub Push(v As Variant)
    ' store passed variant on a "stack" (= collection)
    SafeGetApplicationStates.Add v
End Sub

Private Function Pop() As Variant
    With SafeGetApplicationStates
        If .Count <> 0 Then
            ' return the highest element = the last one pushed to the "stack"
            Pop = .Item(.Count)
            
            ' remove this item
            .Remove (.Count)
        End If
    End With
End Function

Function PopStatusBar()
    Dim vStatusBar: vStatusBar = Pop()
    
    ' Application.StatusBar is a weird thing, because it is both a boolean (if we set it, to clear it) or a string
    ' So make sure that we restore a cleared status bar to a clear status bar, not a status bar reading "FALSCH"
    If vStatusBar = "FALSE" Or vStatusBar = "FALSCH" Then
        PopStatusBar = False
    Else
        PopStatusBar = vStatusBar
    End If
End Function



'>> Clipboard

' ((VROFZRA))refactor PutOnClipboard, ausgliedern in Objekt

Sub PutOnClipboard(s As String)
    If s <> "" Then
        Dim DataObj As New MSForms.DataObject
        DataObj.SetText s
        DataObj.PutInClipboard
    End If
End Sub

Function GetFromClipboard() As String
    Dim DataObj As New MSForms.DataObject
    DataObj.GetFromClipboard
    On Error Resume Next
    GetFromClipboard = DataObj.GetText(1)
End Function



' >> Colors

'*************************************
' helpers
'*************************************

Sub ClearInteriorColor(r As Range)
    With r.Interior
        .PATTERN = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Function BLACK() As Long
    BLACK = RGB(0, 0, 0)
End Function

Function LIGHT_GRAY() As Long
    LIGHT_GRAY = RGB(245, 245, 245)
End Function



'>> Comments

Sub ShowCommentForm(Optional sCaption = "")
    Dim frm As New UserFormComment
    If sCaption <> "" Then frm.Caption = sCaption
    frm.Show
End Sub


'*************************************
' dumping comments
'*************************************

Sub DumpAllCommentsForAllSheetsOnCommentsSheet(wsComment As Worksheet)
    SetSilentApplicationState
    wsComment.Activate
    Dim rActive As Range: Set rActive = ActiveCell
    Dim rFirst As Range: Set rFirst = Range("A2")
    Dim rLast As Range: Set rLast = rFirst.End(xlDown)
    Dim rColsToClear As Range: Set rColsToClear = wsComment.Columns("A:B")
    Dim rRowsToClear As Range: Set rRowsToClear = wsComment.Range(rFirst, rLast)
    Intersect(rRowsToClear, rColsToClear).ClearContents
    rFirst.Select
    DumpAllCommentsForAllSheets
    ActiveSheet.Cells.EntireRow.AutoFit
    rActive.Select
    RevertApplicationState
End Sub

Sub Button_DumpAllCommentsForAllSheetsDefault()
    DumpAllCommentsForAllSheetsOnCommentsSheet Worksheets(WS_COMMENTS_DEFAULT)
End Sub

Sub DumpAllCommentsForAllSheets()
    ' called from AHK
    
    SetSilentApplicationState
    On Error GoTo done
    
    Dim rNoCommentsWorksheets As Range: Set rNoCommentsWorksheets = SafeGetNamedRange(R_NoCommentsWorksheets)
    Dim dictNoCommentsWorksheets As New Dictionary
    If Not rNoCommentsWorksheets Is Nothing Then
        Dim rNoCommentsWorksheet As Range
        For Each rNoCommentsWorksheet In rNoCommentsWorksheets
            Dim sNoCommentsWorksheet As String: sNoCommentsWorksheet = Trim(rNoCommentsWorksheet)
            If sNoCommentsWorksheet <> "" Then
                dictNoCommentsWorksheets.Add sNoCommentsWorksheet, sNoCommentsWorksheet
            End If
        Next
    End If
    
    Dim rAnchor As Range: Set rAnchor = ActiveCell
    Dim ws As Worksheet
    For Each ws In Worksheets
        If Not dictNoCommentsWorksheets.Exists(ws.Name) Then
            Dim nComments As Integer
            nComments = DumpAllCommentsForSheet(ws, rAnchor)
            Set rAnchor = rAnchor.Offset(nComments)
        End If
    Next
    
done:
    RevertApplicationState
End Sub

Function DumpAllCommentsForSheet(ws As Worksheet, rAnchor As Range, Optional bExternal As Boolean = True) As Integer
    ' assume we want to dump the comments of this sheet
    ' dump (2 columns) begins in active cell
    
    Dim nComments As Integer
    
    Dim cmt As Comment
    For Each cmt In ws.Comments
        Dim rCommentCell As Range: Set rCommentCell = cmt.Parent
        
        nComments = nComments + 1
        
        ' col 1: show address of comment (can be used for Ctrl-P)
        rAnchor.Cells(nComments, 1).Formula = "=GetAddress(" & rCommentCell.Address(External:=True) & ", " & IIf(bExternal, "true", "false") & ")"
        
        ' col 1: show comment text
        rAnchor.Cells(nComments, 2).Formula = "=GetComment(" & rAnchor.Cells(nComments, 1).Address & ", true)"
        
    Next
    
    DumpAllCommentsForSheet = nComments
End Function

'*************************************
' getters
'*************************************

Function GetComment(r As Range, Optional bIndirect As Boolean) As Variant
    'Application.Volatile

    GetComment = ""
    Dim sComment As String

    On Error GoTo done
    If bIndirect Then
        Dim sAddress As String: sAddress = r.Value2
        sComment = Range(sAddress).Comment.Text
    Else
        sComment = r.Comment.Text
    End If
    
    ' we'll show the comment either directly in the cell w/ GetComment() or collect the comment and display it on a comments sheet
    ' PROBLEM: Excel doesn't allow length of string input in cell > 254
    ' 28.02.21 Das Problem scheint nicht mehr zu existeren, daher Spezialbehandlung in der nächsten Zeile rausgenommen
    ' If Len(sComment) > 253 Then sComment = Left(sComment, 253) & "…"
    
    GetComment = sComment
done:
End Function



'>> Files

'*************************************
' save string to file
'*************************************

Sub SaveStringToTextFile(s As String, sfn As String)
    Open sfn For Output As #1
    Print #1, s
    Close #1
End Sub

Sub SaveCollectionToTextFile(c As Collection, sfn As String)
    Dim s As String: s = CollectionToString(c, Chr(10))
    SaveStringToTextFile s, sfn
End Sub



'>> Macros

Sub CtrlI()
Attribute CtrlI.VB_ProcData.VB_Invoke_Func = "i\n14"
    ShowCommentForm
End Sub

Sub CtrlJ()
Attribute CtrlJ.VB_ProcData.VB_Invoke_Func = "j\n14"
    GotoPrevSection
End Sub

Sub CtrlK()
Attribute CtrlK.VB_ProcData.VB_Invoke_Func = "k\n14"
    GotoNextSection
End Sub

Sub CtrlR()
Attribute CtrlR.VB_ProcData.VB_Invoke_Func = "r\n14"
    RenameSheet
End Sub



'>> Names

'*************************************
' getters
'*************************************

Function NamedRangeExists(sRangeName As String, Optional ws As Worksheet) As Boolean
    NamedRangeExists = False
    On Error GoTo done
    
    Dim r As Range
    If ws Is Nothing Then
        ' global
        Set r = Range(sRangeName)
    Else
        ' sheet-local
        Set r = ws.Range(sRangeName)
    End If
    NamedRangeExists = True
done:
End Function

Function SafeGetNamedRange(sRangeName As String, Optional ws As Worksheet) As Range
    Set SafeGetNamedRange = Nothing
    On Error GoTo done
    
    If ws Is Nothing Then
        ' global
        Set SafeGetNamedRange = Range(sRangeName)
    Else
        ' sheet-local
        Set SafeGetNamedRange = ws.Range(sRangeName)
    End If
done:
End Function



'>> Params

Function DoFormatSentence2(sSentence As String, arrParams() As Variant) As String
    sSentence = Replace(sSentence, "\#", "@@@")
    
    Dim ixParam As Integer
    Dim ix As Integer
    Do
        ix = InStr(1, sSentence, "#")
        Dim bVariableFound As Boolean: bVariableFound = ix <> 0
        If bVariableFound Then
            
            If ixParam > UBound(arrParams) Then
            
                DoFormatSentence2 = "# missing parameters"
                Exit Function
            
            End If
            
            ' no special handling of Range needed, because we stringified the array
            Dim sReplacement As String: sReplacement = arrParams(ixParam)
        
            sSentence = Left(sSentence, ix - 1) & sReplacement & Mid(sSentence, ix + 1)
        
            ' different than above, because we directly work with 0-based array coming from a ParamArray
            ixParam = ixParam + 1
        End If
    Loop While bVariableFound
    
    sSentence = Replace(sSentence, "@@@", "#")
    
    DoFormatSentence2 = sSentence
End Function

Function FS(sSentence As String, ParamArray tokens() As Variant) As String
    Dim arrParams(): arrParams = tokens
    FS = DoFormatSentence2(sSentence, arrParams)
End Function


'*************************************
' lowlevel helpers
'*************************************

Public Function ParamArrayDelegated(ParamArray prms() As Variant) As Variant
    ' http://stackoverflow.com/questions/20783170/pass-array-to-paramarray
    
    Dim arrPrms() As Variant, arrWrk() As Variant
    'When prms(0) is Array, supposed is delegated from another function
    arrPrms = prms
    Do While VarType(arrPrms(0)) >= vbArray And Not IsObject(arrPrms(0)) And UBound(arrPrms) < 1
        arrWrk = arrPrms(0)
        arrPrms = arrWrk
    Loop
    ParamArrayDelegated = arrPrms
End Function



'>> RangeGetters

Function LastInColumn2(rColumn As Range) As Range
    Dim rSourceTop As Range: Set rSourceTop = rColumn.Cells(1)
    Dim rLastInColumn As Range: Set rLastInColumn = rColumn.Cells(rColumn.Cells.Count)
    Dim rSourceEndUp As Range: Set rSourceEndUp = rLastInColumn.End(xlUp)
    Set LastInColumn2 = rSourceEndUp
End Function



'>> RegExpr

'*************************************
' helpers
'*************************************

Function CreateRegExp(sSubstring As String, Optional bCaseSensitive As Boolean, Optional bMultiLine As Boolean, Optional bGlobal As Boolean = True) As RegExp
    Set CreateRegExp = New RegExp
    With CreateRegExp
        .Global = bGlobal
        .MultiLine = bMultiLine
        .IgnoreCase = Not bCaseSensitive
        .PATTERN = sSubstring
    End With
End Function

Function SafeCreateRegExp(sPattern As String, bCaseSensitive As Boolean, bMultiLine As Boolean, Optional bGlobal As Boolean = True) As RegExp
    ' make sure the regex buffer exists
    If g_dRegExps Is Nothing Then Set g_dRegExps = New Dictionary
    
    Dim RE As RegExp
    
    ' get the regex, ...
    If g_dRegExps.Exists(sPattern) Then
    
        ' ... either from buffer ...
        Set RE = g_dRegExps(sPattern)
    Else
    
        ' ... or create new and store in buffer
        Set RE = CreateRegExp(sPattern, bCaseSensitive, bMultiLine, bGlobal)
        g_dRegExps.Add sPattern, RE
    End If


    Set SafeCreateRegExp = RE
End Function

Function RegMatch(vTarget, sPattern As String, Optional ixGroup, Optional bCaseSensitive As Boolean, Optional bMultiLine As Boolean, Optional bGlobal As Boolean = True) As Variant
    If IsMissing(ixGroup) Then ixGroup = -1
    
    Dim sTarget As String: sTarget = vTarget
    Dim RE As RegExp: Set RE = SafeCreateRegExp(sPattern, bCaseSensitive, bMultiLine, bGlobal)
    
    If ixGroup = -1 Then
    
        ' nothing passed -> return true if it matches
        RegMatch = RE.Test(sTarget)
        
    Else
    
        Dim matches As Object
    
        If vTarget = m_LastTarget And sPattern = m_LastPattern Then
            ' reuse last, if we use multiple group indexes on the same target/pattern
            Set matches = m_LastMatches
        Else
            ' extract group match
            Set matches = RE.Execute(sTarget)
        End If
        Set m_LastMatches = matches
        m_LastTarget = vTarget
        m_LastPattern = sPattern
        
        Dim sMatch As String
        On Error Resume Next
        If ixGroup = 0 Then
            sMatch = matches(0).value
        Else
            sMatch = matches(0).SubMatches.Item(ixGroup - 1)
        End If
        LastMatch0 = sMatch
        RegMatch = sMatch
    End If
End Function



'>> SheetJump

'*************************************
' show / hide sheets in section
'*************************************

Sub Button_ToggleSheetsInSection()
    Dim btn As Shape: Set btn = ActiveSheet.Shapes(Application.Caller)
    HideSheetsInSection btn
End Sub

Function GetHideOrShowButton(ws As Worksheet)
    Dim btn As Button
    For Each btn In ws.Buttons
        If btn.Caption = "hide section sheets" Or btn.Caption = "show section sheets" Then
            Set GetHideOrShowButton = btn
            Exit Function
        End If
    Next
    Set GetHideOrShowButton = Nothing
End Function

Sub HideOrShowSheetsInAllSections(bShow As Boolean)
    SetSilentApplicationState
    Dim old: Set old = ActiveSheet
    Dim ws As Worksheet
    For Each ws In Worksheets
        If IsSectionSheet(ws) Then
            HideOrShowSheetsInSection ws, bShow
            Dim btn As Button: Set btn = GetHideOrShowButton(ws)
            If Not btn Is Nothing Then
                Dim sNewCaption As String
                If bShow Then sNewCaption = "hide section sheets" Else sNewCaption = "show section sheets"
                btn.Caption = sNewCaption
            End If
        End If
    Next
    old.Activate
    RevertApplicationState
End Sub

' was ShowSheetsInAllSections

Sub Button_ShowSheetsInAllSections()
    HideOrShowSheetsInAllSections True
End Sub

' was HideSheetsInAllSections

Sub Button_HideSheetsInAllSections()
    HideOrShowSheetsInAllSections False
End Sub

Sub HideOrShowSheetsInSection(wsSection As Worksheet, bShow As Boolean)
    If Not IsSectionSheet(wsSection, False) Then Exit Sub
    Dim ws As Worksheet: Set ws = wsSection
    Do
        Set ws = ws.Next
        If ws Is Nothing Then Exit Do
        If IsSectionSheet(ws, False) Or ws.index = Sheets.Count Then Exit Do
        ws.Visible = bShow
    Loop Until IsSectionSheet(ws, False)
End Sub

Sub HideSheetsInSection(btn As Shape)
    If Not IsSectionSheet(ActiveSheet, False) Then Exit Sub
    
    SetSilentApplicationState
    
    Dim wsActive As Worksheet: Set wsActive = ActiveSheet
    Dim bShow As Boolean: bShow = btn.TextFrame.Characters.Caption = "show section sheets"
    
    HideOrShowSheetsInSection wsActive, bShow
    
    Dim sNewCaption As String
    If bShow Then sNewCaption = "hide section sheets" Else sNewCaption = "show section sheets"
    btn.TextFrame.Characters.Caption = sNewCaption
    
    wsActive.Activate
    
    RevertApplicationState
End Sub

Function IsSectionSheet(ws As Worksheet, Optional bIncludeArrowSheets As Boolean = True) As Boolean
    With ws.Tab
        If .Color = RGB(43, 22, 75) Or (bIncludeArrowSheets And (StartsWith(ws.Name, ">") Or StartsWith(ws.Name, "<") Or EndsWith(ws.Name, ">"))) Then
            IsSectionSheet = True
        Else
            If .Color <> RGB(0, 0, 0) Then
                IsSectionSheet = False
            Else
                IsSectionSheet = .ColorIndex <> xlColorIndexNone
            End If
        End If
    End With
End Function

Sub CreateSectionAgendas()

    SetSilentApplicationState
    Dim wsOld: Set wsOld = ActiveSheet

    Dim cSections As New Collection
    Dim ws As Worksheet
    For Each ws In Worksheets
        If IsSectionSheet(ws, False) Then cSections.Add ws.Name
    Next
    
    Dim vSections: vSections = Application.WorksheetFunction.Transpose(CollectionToArray(cSections))
    
    Dim iSectionMidpoint As Integer
    If NamedRangeExists(R_SectionMidpoint) Then
        iSectionMidpoint = Range(R_SectionMidpoint)
    Else
        iSectionMidpoint = DEFAULT_SECTION_MIDPOINT
    End If
    
    Dim ixSectionSheet As Integer
    For Each ws In Worksheets
        
        If IsSectionSheet(ws, False) Then
            ixSectionSheet = ixSectionSheet + 1
        
            ws.Columns("G").ClearContents
            Dim rMidpoint As Range: Set rMidpoint = ws.Range("G" & iSectionMidpoint)
            
            Dim rAnchor As Range: Set rAnchor = ws.Range(rMidpoint.Offset(-ixSectionSheet + 1), rMidpoint.Offset(cSections.Count - ixSectionSheet))
            
            rAnchor.Value2 = vSections
            rAnchor.Font.Size = 8
            rAnchor.Font.Bold = False
        
            rMidpoint.Font.Size = 12
            rMidpoint.Font.Bold = True
            
            ws.Activate
            ActiveWindow.DisplayHeadings = False
        End If
    Next
    
    wsOld.Activate
    RevertApplicationState

End Sub

Sub GotoNextSection()
    On Error GoTo done
    
    Dim ixSheet As Integer
    
    ' find next section sheet; +1: we might start looking while being on a section sheet
    For ixSheet = ActiveSheet.index + 1 To Worksheets.Count
        Dim ws As Worksheet: Set ws = Worksheets(ixSheet)
        
        ' check if Registerfarbe indicates that it is a section sheet
        If IsSectionSheet(ws, True) Or IsPartSheet(ws) Then
            ws.Activate
            Exit Sub
        End If
        
    Next ixSheet
done:
End Sub

Sub GotoPrevSection()
    On Error GoTo done
    
    Dim ixSheet As Integer
    
    ' find next section sheet; -1: see GotoNextSection
    For ixSheet = ActiveSheet.index - 1 To 1 Step -1
        Dim ws As Worksheet: Set ws = Worksheets(ixSheet)
        
        ' check if Registerfarbe indicates that it is a section sheet
        If (IsSectionSheet(ws, True) Or IsPartSheet(ws)) And ws.Visible Then
            ws.Activate
            Exit Sub
        End If
        
    Next ixSheet
done:
End Sub


'*************************************
' parts (bigger steps) - ^n, ^m
'*************************************

Function IsPartSheet(ws As Worksheet) As Boolean
    IsPartSheet = ws.Tab.Color = RGB(13, 13, 13)
End Function



'>> Sheets

Public Function GetWorksheetName(r As Range) As String
    GetWorksheetName = r.Worksheet.Name
End Function

Function SheetExists(sName, Optional wb As Workbook) As Boolean
    On Error Resume Next
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim wsExisting As Worksheet: Set wsExisting = wb.Worksheets(sName)
    SheetExists = Not (wsExisting Is Nothing)
End Function


'*************************************
' helpers
'*************************************

Sub RenameSheet(Optional ws As Worksheet, Optional sNewName As String)
    If ws Is Nothing Then Set ws = ActiveSheet
    If sNewName = "" Then sNewName = InputBox("new sheet name:", "rename sheet", ws.Name)
    sNewName = Trim(sNewName)
    Dim bWantRename As Boolean: bWantRename = sNewName <> "" And ws.Name <> sNewName
    If bWantRename Then
        On Error Resume Next
        Dim wsExisting As Worksheet: Set wsExisting = Worksheets(sNewName)
        On Error GoTo 0
        If Not SheetExists(sNewName) Then
            ws.Name = sNewName
        Else
            MsgBox "Sheet '" & sNewName & "' already exists"
        End If
    End If
End Sub



'>> StatusBar

Public Sub ClearStatusBar()
    Application.StatusBar = False
    DoEvents
End Sub

Public Sub ShowStatus(v As Variant)
    DoEvents
    Application.StatusBar = v
    DoEvents
End Sub



'>> Strings

Function AsString(v As Variant) As String
    If IsError(v) Then
        AsString = "#NV"
    Else
        On Error Resume Next
        AsString = v
    End If
End Function


'*************************************
' string <-> collection
'*************************************

Public Function StringToCollection(s As String, Optional sDelim As String) As Collection
    If sDelim = "" Then sDelim = "|"
    s = Trim(s)
    Set StringToCollection = New Collection
    If s = "" Then Exit Function
    
    Dim aTags: aTags = Split(s, sDelim)
    Dim ixTag As Integer
    For ixTag = LBound(aTags) To UBound(aTags)
        Dim sTag As String: sTag = Trim(aTags(ixTag))
        StringToCollection.Add sTag
    Next ixTag
End Function

Public Function CollectionToString(c As Collection, Optional sDelim As String) As String
    If sDelim = "" Then sDelim = "|"
    Dim s As String
    Dim bFirst As Boolean: bFirst = True
    Dim v
    For Each v In c
        If bFirst Then
            bFirst = False
            s = v
        Else
            s = s & sDelim & v
        End If
    Next
    CollectionToString = s
End Function


'*************************************
' tests
'*************************************

Function StartsWith(sTarget As String, sCheck As String, Optional bCaseSensitive As Boolean) As Boolean
    If bCaseSensitive Then
        StartsWith = Left(sTarget, Len(sCheck)) = sCheck
    Else
        ' 15.7. problem: sCheck=UCase(sChec) changes passed sCheck outside the function (even w/ ByVal)
        StartsWith = Left(UCase(sTarget), Len(sCheck)) = UCase(sCheck)
    End If
End Function

Function EndsWith(sTarget As String, sCheck As String, Optional bCaseSensitive As Boolean) As Boolean
    If bCaseSensitive Then
        EndsWith = Right(sTarget, Len(sCheck)) = sCheck
    Else
        EndsWith = Right(UCase(sTarget), Len(sCheck)) = UCase(sCheck)
    End If
End Function



'>> Util2

Function CollectionToArray(c As Collection, Optional b2dim As Boolean = False) As Variant()
    Dim av()
    If b2dim Then ReDim av(1 To c.Count, 1 To 1) Else ReDim av(1 To c.Count)
    Dim ix As Long: ix = 1
    Dim v
    For Each v In c
        If IsObject(v) Then
            If b2dim Then Set av(ix, 1) = v Else Set av(ix) = v
        Else
            If b2dim Then av(ix, 1) = v Else av(ix) = v
        End If
        ix = ix + 1
    Next
    CollectionToArray = av
End Function



' >> Version

Sub SetVersion()
    Dim sVersion As String: sVersion = "Version: " & Format(Now, "YYMMDD vHHMM")
    sVersion = InputBox("Version", "Version", sVersion)
    If sVersion = "" Then Exit Sub
    
    Dim vRes: vRes = MsgBox("set version '" & sVersion & "' for this sheet only?", vbYesNoCancel)
    If vRes = vbCancel Then Exit Sub
    
    Dim sVersionDate As String: sVersionDate = RegMatch(sVersion, "([0-9]+ v[0-9]+)$", 1)
    PutOnClipboard sVersionDate
    
    If vRes = vbYes Then
        SetVersionBox ActiveSheet, sVersion
    Else
        vRes = MsgBox("set version '" & sVersion & "' for ALL sheet?s", vbYesNoCancel)
        If vRes = vbYes Then
            Dim ws As Worksheet
            For Each ws In Worksheets
                SetVersionBox ws, sVersion
            Next
        End If
    End If
End Sub

Sub SetVersionBox(ws As Worksheet, sVersion As String)
    On Error Resume Next
    'If ws.Name = "Droplist" Then Stop
    ws.Shapes.Range(R_VersionBox).TextFrame2.textRange.Characters.Text = sVersion
End Sub
