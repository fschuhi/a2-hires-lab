Attribute VB_Name = "A_"
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

' 05.01.25 fix ShowStatus multiple DoEvens fix
' 06.02.25 added ReplaceSpaces, bitmaps for Powerpoint
' 17.02.25 FMonth
' 15.03.25 Excel SlideText helpers
' 16.03.25 SimulateSlideText, SafeAdd, moved JumpStation mechanics to here, FPctP
' 29.03.25 ColorSelection
' 21.05.26 QuickSortStrings
' 01.06.26 fixed multiple CASE in ColorFormulas
' 01.07.26 added GetRangeName, FirstNumberIndex, LastNumberIndex from model+BIB

Option Explicit

Const TODO_ID = 1

' ApplicationStates

' prevent Worksheet_Change
Public WorksheetChangeDisabled As Boolean

Private g_cApplicationStates As Collection


' Comments

Const WS_COMMENTS_DEFAULT = "> comments <"

Public g_iUserFormCommentTop As Integer
Public g_iUserFormCommentLeft As Integer


' Dents

Public Const USERFORMDENTS_LEFT = 900
Public g_UserFormDentsLeft As Integer
Public Const DENTS_TIMEOUT_SECONDS = 5


' Files

Private g_Substs As Dictionary


' JumpStation

Public m_frmSelector As UserFormSelector
Public m_jumpTargets As Object  ' Dictionary<String, String>
Public m_sheetSpecificTargets As Object  ' Dictionary<String, String>
Public m_cBackSheets As Collection


' Names

Public Const LOCAL_NAME = True
Public Const GLOBAL_NAME = False
Public Const DO_NOT_MARK_NAMED_RANGE = True
Public Const MARK_NAMED_RANGE = False
Public Const R_NoCommentsWorksheets = "NoCommentsWorksheets"


' Pomodoros

Const PAT_LineParts = "(.+) \(([0-9]+)\)"


' Powerpoint

Const DO_NOT_RETAIN_WIDTH = False
Const RETAIN_WIDTH = True
Public Const R_SLIDE_TEXTS = "SlideTexts"

' while debugging and adding fucntionality
'Private g_Powerpoint As PowerPoint.Application

' in production - - easier w/ Verweise
Private g_Powerpoint As Object


' RangeGetters

Public Const xlOrientationCell = 0
Public Const xlOrientationRow = 1
Public Const xlOrientationColumn = 2
Public Const xlOrientationMatrix = 3

Private m_RangeObjects As New Dictionary


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


' Sheets

Public Const EXCEL_LAST_ROW_PJ = 1048574
Public Const EXCEL_LAST_ROW = 1048576
Public Const EXCEL_LAST_COL = 16384


' Strand

Public Const PAT_Strand = "\[[A-Za-z][A-Za-z][A-Za-z][A-Za-z][A-Za-z][A-Za-z][A-Za-z]?\]"
Public Const PAT_Strand_Line = "^(\[[A-Za-z][A-Za-z][A-Za-z][A-Za-z][A-Za-z][A-Za-z][A-Za-z]?\])(.+)?$"


' Speech

Public Const VOICE_DE = 0
Public Const VOICE_EN_M = 1
Public Const VOICE_EN_F = 2


' Timer

Private g_dTimer As Double


' Timestamps

Const R_TIMESTAMP = "Timestamp"


' Util2

Public Const PERCENTAGE_TOO_BIG As Double = 3
Public Const EPSILON As Double = 0.000001

' Versions

Const R_VersionBox = "VersionBox"



#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
#End If

Public g_TEMP As String


#If VBA7 Then
    Public Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
#Else
    Public Declare Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
#End If
#If VBA7 Then
    Public Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
    Public Declare Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If


#If VBA7 Then
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If


' System

Public Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long



'>> Addresses

'*************************************
' Address stuff
'*************************************

Public Function GetFormula(r As Range) As String
    GetFormula = r.formula
End Function

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


Public Function GetAddress2(r As Range) As String
    GetAddress2 = r.Address
End Function


Function GetAddress3(r As Range) As String
    Dim sFullAddress As String: sFullAddress = r.Address(External:=True)
    
    Dim sWorkbook As String: sWorkbook = RegMatch(sFullAddress, "^'(\[.+?\])([^']+?)'!(.+$)", 1)
    Dim sSheet As String: sSheet = RegMatch(sFullAddress, "^'(\[.+?\])([^']+?)'!(.+$)", 2)
    Dim sLocalAddress As String: sLocalAddress = RegMatch(sFullAddress, "^'(\[.+?\])([^']+?)'!(.+$)", 3)
    
    Dim sAddress
    If sSheet <> Application.Caller.Worksheet.Name Then sAddress = "'" & sSheet & "'!"
    sAddress = sAddress & sLocalAddress
    
    GetAddress3 = sAddress
End Function


Public Function GetAddressVolatile(r As Range, Optional bWithSheet As Boolean) As String
    Application.Volatile
    GetAddressVolatile = GetAddress(r, bWithSheet)
End Function


Function TrimCellAddressOld(sAddress) As String
    sAddress = Replace(sAddress, "[" & ActiveWorkbook.Name & "]", "")
    sAddress = Replace(sAddress, "'" & ActiveSheet.Name & "'!", "")
    sAddress = Replace(sAddress, ActiveSheet.Name & "!", "")
    TrimCellAddressOld = sAddress
End Function


Function TrimCellAddress(sAddress) As String
    Dim sNewAddress As String: sNewAddress = sAddress
    sNewAddress = Replace(sNewAddress, "[" & ActiveWorkbook.Name & "]", "")
    Dim bOwnWorkbook As Boolean: bOwnWorkbook = sNewAddress <> sAddress
    If bOwnWorkbook Then
        sNewAddress = Replace(sNewAddress, "'" & ActiveSheet.Name & "'!", "")
        sNewAddress = Replace(sNewAddress, ActiveSheet.Name & "!", "")
    End If
    TrimCellAddress = sNewAddress
End Function


'*************************************
' address manipulation
'*************************************

Sub TestModifyFormulas()
    Dim s As String: s = MakeFormulaAbsolute("INDEX($BV$11:$BX$11,MATCH($AU$7,$BV$7:$BX$7,0))", ActiveSheet)
    s = MakeFormulaAbsolute(s, ActiveSheet)
    s = MakeFormulaAbsolute(s, ActiveSheet)
    s = MakeFormulaAbsolute(s, ActiveSheet)
    Debug.Assert s = "INDEX('scenarios'!$BV$11:$BX$11,MATCH('scenarios'!$AU$7,'scenarios'!$BV$7:$BX$7,0))"
End Sub


Function RemoveWorkbookFromFormula(sFormula As String) As String
    Dim sRemaining As String: sRemaining = sFormula
    Dim sNewFormula As String: sNewFormula = ""
    Dim sMatch As String
    Dim sSheet As String
    Dim ixFound As String
    Do
        sMatch = RegMatch(sRemaining, "'?(\[[^[]+\]([^!']+))!", 1)
        If sMatch <> "" Then
            ixFound = InStr(1, sRemaining, sMatch)
            sSheet = RegMatch(sFormula, "'?(\[[^[]+\]([^!']+))!", 2)
            sNewFormula = sNewFormula & Left(sRemaining, ixFound - 1) & "'" & sSheet & "'"
            sRemaining = Mid(sRemaining, ixFound + Len(sMatch))
        End If
    Loop Until sRemaining = "" Or sMatch = ""
    sNewFormula = sNewFormula & sRemaining
    RemoveWorkbookFromFormula = sNewFormula
End Function


Private Function ConvertFormula(sFormula As String) As String
    On Error GoTo err
    ConvertFormula = Application.ConvertFormula(sFormula, xlA1, xlA1, xlAbsolute)
    Exit Function
err:
    ConvertFormula = sFormula
End Function

Function MakeFormulaAbsolute(sFormula As String, ws As Worksheet) As String
    sFormula = ConvertFormula(sFormula)
    sFormula = RemoveWorkbookFromFormula(sFormula)
    
    Dim sNewFormula As String
    
    Dim sRemaining As String
    Dim sMatch As String
    Dim ixFound As Integer
    
    sRemaining = sFormula
    sNewFormula = ""
    Do
        sMatch = Mid(RegMatch(sRemaining, "([^!:]\$[A-Z][A-Z]?\$[0-9]+:\$[A-Z][A-Z]?\$[0-9]+)", 1), 2)
        If sMatch <> "" Then
            ixFound = InStr(1, sRemaining, sMatch)
            sNewFormula = sNewFormula & Left(sRemaining, ixFound - 1) & "'" & ws.Name & "'!" & sMatch
            sRemaining = Mid(sRemaining, ixFound + Len(sMatch))
        End If
    Loop Until sRemaining = "" Or sMatch = ""
    sNewFormula = sNewFormula & sRemaining
    
    sRemaining = sNewFormula
    sNewFormula = ""
    Do
        sMatch = Mid(RegMatch(sRemaining, "([^!:]\$[A-Z][A-Z]?\$[0-9]+)", 1), 2)
        If sMatch <> "" Then
            ixFound = InStr(1, sRemaining, sMatch)
            sNewFormula = sNewFormula & Left(sRemaining, ixFound - 1) & "'" & ws.Name & "'!" & sMatch
            sRemaining = Mid(sRemaining, ixFound + Len(sMatch))
        End If
    Loop Until sRemaining = "" Or sMatch = ""
    sNewFormula = sNewFormula & sRemaining
    
    MakeFormulaAbsolute = sNewFormula
End Function


Function HasFormula(r As Range, Optional sPassedFormulaName As String) As Boolean
    ' only works for single cells
    Debug.Assert r.Cells.Count = 1
    
    If Not r.HasFormula Then
    
        ' trivially, the cells doesn't have a formula if it doesn't have a formula :)
        HasFormula = False
        
    Else
            
        If sPassedFormulaName = "" Then
        
            ' we just want to know if there is any formula in the cell
            HasFormula = True
        Else
        
            ' not trivial: is the formula the correct one?
            Dim sFormulaName As String: sFormulaName = RegMatch(r.formula, "^=( ?- ?)?([A-z0-9]+)\(", 2)
            HasFormula = sFormulaName = sPassedFormulaName
        End If
    End If
End Function


'*************************************
' address management (abs, rel)
'*************************************

Function GetAddressAbsoluteRelative(r As Range) As String
    Dim sAddress As String: sAddress = Replace(r.Address, "$", "")
    Const RE = "\$?([A-Z]+)\$?([0-9]+)"
    Dim sCol As String: sCol = RegMatch(r.Address, RE, 1)
    Dim sRow As String: sRow = RegMatch(r.Address, RE, 2)
    GetAddressAbsoluteRelative = "$" & sCol & sRow
End Function


Function GetAddressRelativeAbsolute(r As Range) As String
    Dim sAddress As String: sAddress = Replace(r.Address, "$", "")
    Const RE = "\$?([A-Z]+)\$?([0-9]+)"
    Dim sCol As String: sCol = RegMatch(r.Address, RE, 1)
    Dim sRow As String: sRow = RegMatch(r.Address, RE, 2)
    GetAddressRelativeAbsolute = sCol & "$" & sRow
End Function



'>> ApplicationStates

'*************************************
' change event disabling
'*************************************

Sub ResetWorksheetChangeDisabled()
    ' usually called from the outside
    ' allow Worksheet_Change
    WorksheetChangeDisabled = False
End Sub


'*************************************
' push / pop
'*************************************

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


' Like SetSilentApplicationState(), but just cutting of screen updating
Sub SetNoScreenUpdating()
    SaveApplicationState
    Application.ScreenUpdating = False
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



'>> Charts

' 15.03.16 added SetAxisColor

'*************************************
' helper functions
'*************************************

Function GetWorksheetFromChart(c As chart) As Worksheet
    Set GetWorksheetFromChart = c.Parent.Parent
End Function

Function GetShapeFromChart(c As chart) As Shape
    Set GetShapeFromChart = ActiveSheet.Shapes(c.Parent.Name)
End Function


'*************************************
' change appearance to "standard"
'*************************************

' called by AHK
Sub SetChartStandards()
    MoveLegendToBottom ActiveChart
    SetGrayGridlines ActiveChart
    SetChartFontName ActiveChart, "Segoe UI"
    
    SetAxisColor ActiveChart, xlCategory, RGB(166, 166, 166), RGB(221, 221, 221)
    SetAxisColor ActiveChart, xlValue, RGB(166, 166, 166), RGB(221, 221, 221)
    
    ActiveChart.PlotVisibleOnly = False
End Sub


Sub MoveLegendToBottom(c As chart)
    ' legend might have been already removed
    On Error Resume Next
    c.Legend.Position = xlBottom
    c.Legend.Format.TextFrame2.textRange.Font.Fill.ForeColor.RGB = RGB(128, 128, 128)
End Sub

Sub SetGrayGridlines(c As chart)
    With c.Axes(xlValue).MajorGridlines.Format.line
        .ForeColor.RGB = RGB(234, 234, 234)
        .Transparency = 0
        .Visible = msoTrue
    End With
    With c.Axes(xlCategory).MajorGridlines.Format.line
        .ForeColor.RGB = RGB(234, 234, 234)
        .Transparency = 0
        .Visible = msoTrue
    End With
End Sub


Sub SetAxisColor(c As chart, t As XlAxisType, clFont As Long, Optional clTicks As Long)
    c.Axes(t).TickLabels.Font.Color = clFont
    If clTicks = 0 Then clTicks = clFont
    c.Axes(t).Format.line.ForeColor.RGB = clTicks
End Sub


Sub SetAxisFormat(c As chart, t As XlAxisType, sFormat As String)
    c.Axes(t).TickLabels.NumberFormat = sFormat
End Sub


Sub SetChartFontName(c As chart, sFontName As String)
    With c.ChartArea.Format.TextFrame2.textRange.Font
        .NameComplexScript = sFontName
        .NameFarEast = sFontName
        .Name = sFontName
        .Size = 8
    End With
End Sub


'*************************************
' GrayGreenRed
'*************************************
Function SingleGGR(vTY, vLY) As Double()
    ' 1: gray über
    ' 2: green über
    ' 3: red über
    ' 4: gray unter
    ' 5: green unter
    ' 6: red unter

    On Error Resume Next
    Dim TY As Double: TY = vTY
    Dim LY As Double: LY = vLY
    On Error GoTo 0

    Dim dGGR() As Double
    ReDim dGGR(1 To 6)
    If TY >= 0 Then
        If LY >= 0 Then
            If TY >= LY Then ' case 1
                dGGR(1) = LY
                dGGR(2) = TY - LY
            Else ' case 2
                dGGR(1) = TY
                dGGR(3) = LY - TY
            End If
        Else ' case 3
            ' 1 > 2 by definition
            dGGR(2) = TY
            dGGR(4) = LY
        End If
    Else
        If LY >= 0 Then ' case 4
            ' 1 < 2 by definition
            dGGR(1) = LY
            dGGR(6) = TY
        Else
            If TY > LY Then ' case 5
                dGGR(4) = TY
                dGGR(5) = LY - TY
            Else ' case 6
                dGGR(4) = LY
                dGGR(6) = TY - LY
            End If
        End If
    End If
    
    SingleGGR = dGGR
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

Sub PutInClipboard(s As String)
    PutOnClipboard s
End Sub


Sub PutCellContentOnClipboard()
    PutOnClipboard ActiveCell.Value2
End Sub


Function GetFromClipboard() As String
    Dim DataObj As New MSForms.DataObject
    DataObj.GetFromClipboard
    On Error Resume Next
    GetFromClipboard = DataObj.GetText(1)
End Function


' >> Colors

'*************************************
' getters
'*************************************

Function GetCellColor(r As Range) As Long
    GetCellColor = r.Interior.Color
End Function

Function GetFontColor(r As Range) As Long
    GetFontColor = r.Font.Color
End Function


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


'*************************************
' Macros (AHK)
'*************************************

Public Sub SetSelectionFontColorRed()
    Selection.Font.Color = RGB(255, 0, 0)
End Sub

Public Sub SetBackgroundColorLightGreen()
    Selection.Interior.Color = RGB(235, 241, 222)
End Sub


Sub PasteRGBToClipboardForSelectedCell()
    Dim r As Integer: Dim g As Integer: Dim b As Integer
    DetermineRGB Selection, r, g, b
    PutInClipboard VerboseRGB(r, g, b)
End Sub

Sub NoteVerboseInteriorColors()
    Dim r As Range
    For Each r In Selection
        r.value = VerboseRGB(r.Interior.Color)
    Next
End Sub

Sub ColorAccordingToVerboseRGB()
    Const RE = "RGB\(([0-9]+);([0-9]+);([0-9]+)\)"
    Dim rToColor As Range
    For Each rToColor In Selection
        
        If IsNumeric(rToColor) Then
            rToColor.Interior.Color = rToColor.Value2
        Else
            Dim sVerboseRGB As String: sVerboseRGB = rToColor.Value2
            If RegMatch(sVerboseRGB, RE) Then
                Dim r As Integer, g As Integer, b As Integer
                r = RegMatch(sVerboseRGB, RE, 1)
                g = RegMatch(sVerboseRGB, RE, 2)
                b = RegMatch(sVerboseRGB, RE, 3)
                rToColor.Interior.Color = RGBtoLong(r, g, b)
            End If
        End If
    Next
End Sub


'*************************************
' standard colors (old UDFs)
'*************************************

Function GreenRGB() As Long
    GreenRGB = RGB(138, 172, 70)
End Function

Function RedRGB() As Long
    RedRGB = RGB(255, 75, 75)
End Function

Function YellowRGB() As Long
    YellowRGB = RGB(255, 255, 0)
End Function

Function PinkRGB() As Long
    PinkRGB = RGB(230, 184, 183)
End Function

Function LightGreenRGB() As Long
    LightGreenRGB = RGB(196, 215, 155)
End Function

Function LightRedRGB() As Long
    LightRedRGB = RGB(250, 191, 143)
End Function


Function GrayRGB(Optional dGray As Double) As Long
    If dGray <= 0 Then dGray = 0.5
    If dGray > 1 Then dGray = 1
    Dim iGray As Integer
    iGray = 255 * dGray
    GrayRGB = RGB(iGray, iGray, iGray)
End Function


'*************************************
' standard colors
'*************************************

Function RED() As Long
    RED = RGB(255, 0, 0)
End Function

Function LIGHTER_RED() As Long
    ' ((KNILAQS)) probably deprecated
    LIGHTER_RED = RGB(250, 191, 143)
End Function

Function LIGHT_RED() As Long
    LIGHT_RED = RGB(255, 75, 75)
End Function

Function BLACK() As Long
    BLACK = RGB(0, 0, 0)
End Function

Function WHITE() As Long
    WHITE = RGB(255, 255, 255)
End Function

Function PURPLE() As Long
    PURPLE = RGB(128, 0, 128)
End Function

Function YELLOW() As Long
    YELLOW = RGB(255, 255, 0)
End Function
    
Function LIGHT_YELLOW() As Long
    LIGHT_YELLOW = RGB(255, 255, 153)
End Function
    
Function ORANGE() As Long
    ORANGE = RGB(255, 102, 0)
End Function
    
Function DARK_GREEN() As Long
    DARK_GREEN = RGB(0, 128, 0)
End Function

Function LIGHT_GREEN() As Long
    LIGHT_GREEN = RGB(0, 210, 95)
End Function

Function LIGHTER_GREEN() As Long
    LIGHTER_GREEN = RGB(196, 215, 155)
End Function

Function LIGHT_LIGHT_GREEN() As Long
    LIGHT_LIGHT_GREEN = RGB(204, 255, 204)
End Function

Function OLIVE_GREEN() As Long
    OLIVE_GREEN = RGB(138, 172, 70)
End Function


Function GRAY(Optional dGray As Double) As Long
    If dGray <= 0 Then dGray = 0.5
    If dGray > 1 Then dGray = 1
    Dim iGray As Integer
    iGray = 255 * dGray
    GRAY = RGB(iGray, iGray, iGray)
End Function

Function DARK_GRAY() As Long
    DARK_GRAY = RGB(150, 150, 150)
End Function

Function LIGHT_GRAY() As Long
    LIGHT_GRAY = RGB(245, 245, 245)
End Function


'*************************************
' cell coloring for analysis
'*************************************

' http://dmcritchie.mvps.org/excel/colors.htm
    
Sub ColorAllCellsOnSheet()
    Application.ScreenUpdating = False
    ColorConstants
    ColorFormulas
    Application.ScreenUpdating = True
End Sub

Sub ColorSelection()
    Application.ScreenUpdating = False
    ColorConstants Selection
    ColorFormulas Selection
    Application.ScreenUpdating = True
End Sub

Sub ColorConstants(Optional rSelection As Range)
    Dim rConstants As Range
    
    ' https://www.excelcampus.com/vba/find-last-row-column-cell/#specialcode
    On Error Resume Next
    If rSelection Is Nothing Then
        Set rConstants = ActiveSheet.Cells.SpecialCells(xlCellTypeConstants)
    Else
        If rSelection.Cells.Count = 1 Then
            Set rConstants = rSelection
            If rConstants.HasFormula Then Exit Sub
        Else
            Set rConstants = rSelection.SpecialCells(xlCellTypeConstants)
        End If
    End If
    On Error GoTo 0
    
    If rConstants Is Nothing Then Exit Sub
    
    Dim rCell As Range
    For Each rCell In rConstants
        If IsNumeric(rCell.Value2) Then
            rCell.Font.Color = RGB(51, 102, 255)
        Else
        
            If rCell.HasFormula Then
            
                Dim sTarget As String: sTarget = rCell.formula
            
                If RegMatch(sTarget, "=( *[A-Z][A-Z]?[0-9]+ *\+)+") Then
                    Application.StatusBar = sTarget
                End If
            
            End If
        
        End If
    Next
End Sub


Function DetermineFormulaPattern(rSelection As Range) As String
    Dim rCell As Range: Set rCell = rSelection.Cells(1, 1)
    Dim sTarget As String: sTarget = rCell.formula
    
    Const R_WORKBOOK = "(\[[^]]+?\])"
    Const R_SHEET = "([A-Za-zÄÖÜäöüß0-9()+,. &_#-]+)"
    Const R_ADDRESS = "(\$?[A-Z][A-Z]?\$?[0-9]+(\*-1)?)"
    Const R_NAME = "([A-Za-z]+)$"
    
    Dim sQualifiedSheet As String: sQualifiedSheet = "'?" & R_WORKBOOK & R_SHEET & "'?!"
    Dim sLocalSheet As String: sLocalSheet = "'?" & R_SHEET & "'?!"

    sTarget = RegReplace(sTarget, R_NAME, "<NAME>")
    sTarget = RegReplace(sTarget, sQualifiedSheet, "<QSHEET>")
    sTarget = RegReplace(sTarget, sLocalSheet, "<SHEET>")
    sTarget = RegReplace(sTarget, R_ADDRESS, "<ADDRESS>")
    sTarget = RegReplace(sTarget, "^ *=[ -]*", "=")
    
    DetermineFormulaPattern = sTarget
End Function


Sub ColorFormulas(Optional rSelection As Range)
    Application.StatusBar = False

    Dim rFormulas As Range
    
    On Error Resume Next
    If rSelection Is Nothing Then
        Set rFormulas = ActiveSheet.Cells.SpecialCells(xlCellTypeFormulas)
    Else
        If rSelection.Cells.Count = 1 Then
            Set rFormulas = rSelection
        Else
            Set rFormulas = rSelection.SpecialCells(xlCellTypeFormulas)
        End If
    End If
    On Error GoTo 0
    
    If rFormulas Is Nothing Then Exit Sub
    
    Dim rCell As Range
    For Each rCell In rFormulas
        Dim sTarget As String: sTarget = rCell.formula
        
        Dim sFormulaPattern As String
    
        If Contains(sTarget, "VLOOKUP", True) Or Contains(sTarget, "XLOOKUP", True) Or Contains(sTarget, "INDEX", True) Or Contains(sTarget, "MATCH", True) Or Contains(sTarget, "SUMIF", True) Or Contains(sTarget, "SUMPRODUCT", True) Or Contains(sTarget, "COUNTIF", True) Then
            rCell.Font.Color = PURPLE
            sFormulaPattern = Trim(DetermineFormulaPattern(rCell))
            If Contains(sFormulaPattern, "<SHEET><ADDRESS>:<ADDRESS>") Then
                rCell.Interior.Color = LIGHT_GRAY
            End If
        Else
        
            If RegMatch(sTarget, "=[ +-]*SUM\(") Then
                rCell.Font.Color = ORANGE
            Else
                sFormulaPattern = Trim(DetermineFormulaPattern(rCell))
                Select Case sFormulaPattern
                    Case "=<ADDRESS>"
                        rCell.Font.Color = DARK_GRAY
                        If rCell.Interior.Color = LIGHT_GRAY Then ClearInteriorColor rCell
                        
                    Case "=<SHEET><ADDRESS>", "=<SHEET><NAME>"
                        rCell.Font.Color = DARK_GRAY
                        rCell.Interior.Color = LIGHT_GRAY
                        
                    Case "=<QSHEET><ADDRESS>"
                        rCell.Font.Color = DARK_GRAY
                        rCell.Interior.Color = LIGHT_GRAY
                        
                    Case Else
                        If RegMatch(sFormulaPattern, "=<ADDRESS>( *[+-] *<ADDRESS>)+") Then
                            rCell.Font.Color = DARK_GREEN
                            If rCell.Interior.Color = LIGHT_GRAY Then ClearInteriorColor rCell
                        Else
                            If RegMatch(sFormulaPattern, "=(<Q?SHEET>)?<ADDRESS>( *[+-] *(<Q?SHEET>)?<ADDRESS>)+") Then
                                rCell.Font.Color = DARK_GREEN
                                rCell.Interior.Color = LIGHT_GRAY
                            Else
                            
                                ' add more here
                            
                            End If
                        End If
                End Select
            End If
        End If
    Next
End Sub


Sub OldColorFormulas(Optional rSelection As Range)
    Application.StatusBar = False

    Dim rFormulas As Range
    
    On Error Resume Next
    If rSelection Is Nothing Then
        Set rFormulas = ActiveSheet.Cells.SpecialCells(xlCellTypeFormulas)
    Else
        If rSelection.Cells.Count = 1 Then
            Set rFormulas = rSelection
        Else
            Set rFormulas = rSelection.SpecialCells(xlCellTypeFormulas)
        End If
    End If
    On Error GoTo 0
    
    If rFormulas Is Nothing Then Exit Sub
    
    Dim rCell As Range
    For Each rCell In rFormulas
        Dim sTarget As String: sTarget = rCell.formula
    
        If Contains(sTarget, "VLOOKUP", True) Or Contains(sTarget, "INDEX", True) Or Contains(sTarget, "MATCH", True) Or Contains(sTarget, "SUMIF", True) Or Contains(sTarget, "SUMPRODUCT", True) Then
            rCell.Font.Color = PURPLE
        Else
        
            If RegMatch(sTarget, "= *SUM\(") Then
                rCell.Font.Color = ORANGE
            Else
            
                Const R_SHEET = "('?\[[A-Za-zAÖÜäöüß0-9()+,. _-]+\])?([A-Za-zÄÖÜäöüß0-9()+,. &'_-]+!)"
                Const R_ADDRESS = "\$?[A-Z][A-Z]?\$?[0-9]+(\*-1)?"
            
                ' grau macht nur dann Sinn, wenn wir auch externe references zulassen
                If RegMatch(sTarget, "^=[ -]*" & R_SHEET & "?" & R_ADDRESS & "$") Then
                    rCell.Font.Color = DARK_GRAY
                    
                    If RegMatch(sTarget, "^=[ -]*" & R_SHEET) Then
                    
                        ' color simple references from other sheets
                        ' background is light gray
                        rCell.Interior.Color = LIGHT_GRAY
                    Else
                        If rCell.Interior.Color = LIGHT_GRAY Then
                        
                            ' if there initially was a ref to another sheet but now (after e.g. refactoring the sheet) there isn't anymore, then we have to remove the indicator cell color
                            ClearInteriorColor rCell
                        End If
                    
                    End If
                    
                Else
                
                    Dim sPattern As String: sPattern = "^=[ -]*" & R_SHEET & "?" & R_ADDRESS
                    If RegMatch(sTarget, sPattern) Then
                    
                        sTarget = RegReplace(sTarget, "'([A-Za-zÄÖÜäöüß0-9()+,. &'_-]+')!", "SHEET!")
                        sTarget = RegReplace(sTarget, R_ADDRESS, "ADDRESS")
                        If RegMatch(sTarget, "SHEET!ADDRESS( *[+-] *SHEET!ADDRESS)+$") Then
                            rCell.Font.Color = DARK_GREEN
                            rCell.Interior.Color = LIGHT_GRAY
                        Else
                        End If
                    
                    Else
                    
                    
                        ' absichtlich hier keine sheet references, weil diese Formeln nicht "einfach Summen" sind
                        If RegMatch(sTarget, "^= *[A-Z][A-Z]?[0-9]+( *[+-] *[A-Z][A-Z]?[0-9]+)+$") Then
                            rCell.Font.Color = DARK_GREEN
                        Else
                        End If
                    End If
                
                End If
            End If
        End If
    Next
End Sub


'*************************************
' import sheets
'*************************************

Sub ColorAsImportSheet(ws As Worksheet)
    With ws.Cells.Interior
        .PATTERN = xlLightUp
        .PatternColor = RGB(200, 200, 200)
        '.PatternThemeColor = xlThemeColorDark1
        '.Color = 13158600
    End With
End Sub


Sub ColorActiveSheetAsImportSheet()
    ColorAsImportSheet ActiveSheet
End Sub


'*************************************
' RGB helpers
'*************************************

Function RGBtoLong(r As Integer, g As Integer, b As Integer) As Long
    RGBtoLong = RGB(r, g, b)
End Function

Sub DetermineRGB(vColor, ByRef r As Integer, ByRef g As Integer, ByRef b As Integer)
    Dim c As Long
    If TypeOf vColor Is Range Then
        Dim rCell As Range: Set rCell = vColor
        c = rCell.Interior.Color
    Else
        c = vColor
    End If
    r = c Mod 256
    g = c \ 256 Mod 256
    b = c \ 65536 Mod 256
End Sub

Function VerboseRGB(i1 As Integer, Optional i2 = -1, Optional i3) As String
    Dim r As Integer, g As Integer, b As Integer
    If i2 = -1 Then
        Dim c As Integer: c = i1
        DetermineRGB c, r, g, b
    Else
        r = i1
        g = i2
        b = i3
    End If
    VerboseRGB = "RGB(" & r & ";" & g & ";" & b & ")"
End Function

Function RGBtoHex(lColor As Long, Optional bLeadingHash As Boolean = True)
    Dim iRed: iRed = (lColor Mod 256)
    Dim iGreen: iGreen = (lColor \ 256) Mod 256
    Dim iBlue: iBlue = (lColor \ 65536) Mod 256
    Dim sHex As String: sHex = VBA.Right$("00" & VBA.Hex(iRed), 2) & VBA.Right$("00" & VBA.Hex(iGreen), 2) & VBA.Right$("00" & VBA.Hex(iBlue), 2)
    If bLeadingHash Then sHex = "#" & sHex
    RGBtoHex = sHex
End Function


Function fnGetCol(strCol As String) As String
    Dim rVal, gVal, bval As String
    strCol = Right("000000" & Hex(strCol), 6)
    bval = Left(strCol, 2)
    gVal = Mid(strCol, 3, 2)
    rVal = Right(strCol, 2)
    fnGetCol = rVal & gVal & bval
End Function


'*************************************
' stripe range
'*************************************

Private Sub ShadeOrUnshadeRow(rRow As Range, bShade As Boolean, clGray As Long)
    Dim r As Range
    For Each r In rRow.Cells
        If bShade Then
            If r.Interior.PATTERN = xlNone Then r.Interior.Color = clGray
        Else
            If r.Interior.Color = clGray Then r.Interior.PATTERN = xlNone
        End If
    Next
End Sub


Sub StripeSomeRange(rToStripe As Range)
    SetSilentApplicationState
    On Error GoTo done
    
    Dim GRAY As Long: GRAY = RGB(242, 242, 242)
    
    Dim bLastVisibleWasShaded As Boolean
    Dim bShade As Boolean
    
    Dim rRow As Range
    For Each rRow In rToStripe.Rows
        If rRow.Hidden Then
            ShadeOrUnshadeRow rRow, False, GRAY
        Else
            bShade = Not bLastVisibleWasShaded
            ShadeOrUnshadeRow rRow, bShade, GRAY
            bLastVisibleWasShaded = bShade
        End If
    Next
done:
    RevertApplicationState
End Sub


Sub StripeRange()
    If Selection.Rows.Count = 1 And Selection.Columns.Count = 1 Then Exit Sub
    StripeSomeRange Selection
End Sub

Sub StripeStripeRange()
    If NameExists("StripeRange") Then
        StripeSomeRange Range("StripeRange")
    Else
        If NameExists("StripedRange") Then
            StripeSomeRange Range("StripedRange")
        Else
            StripeRange
        End If
    End If
End Sub


'*************************************
' range coloring
'*************************************

Sub ColorWithColorCol(ws As Worksheet, Optional rightColor As Long)
    If rightColor = 0 Then rightColor = YELLOW()
    
    Dim rAddressColumn As Range
    On Error Resume Next
    Set rAddressColumn = ws.Range("ColorCol")
    On Error GoTo 0
    Dim bColorColNotDefined As Boolean: bColorColNotDefined = rAddressColumn Is Nothing
    If bColorColNotDefined Then
        MsgBox "define 'ColorCol' (local to this sheet)"
        Exit Sub
    End If
    
    Dim r As Range
    Dim n As Integer: n = Application.WorksheetFunction.CountA(rAddressColumn)
    Dim ix As Integer
    For Each r In rAddressColumn.Cells
        If Not IsEmpty(r) Then

            ColorWithAddressAnchor r, rightColor
            
            ix = ix + 1
            If ix >= n Then Exit For
        End If
    Next
End Sub


Sub ColorWithAddressAnchor(r As Range, rightColor As Long)
    If Not IsEmpty(r) Then
        Dim bIsAddress As Boolean: bIsAddress = Left(r.value, 1) = "$"
        
        If bIsAddress Then
        
            Dim usedRightColor As Long
            Dim rColorId As Range: Set rColorId = r.Offset(0, 1)
            If IsEmpty(rColorId) Then
                usedRightColor = rightColor
            Else
                usedRightColor = rColorId.value
            End If
            
            Dim rDeltas As Range
            'Set rDeltas = ws.Range(r.Offset(1, 0), r.End(xlDown).End(xlToRight))
            Set rDeltas = ExpandedRange(r.Offset(1, 0))
            Dim rTargetTopLeft As Range: Set rTargetTopLeft = r.Worksheet.Range(r.value)
            ColorTargetRange rTargetTopLeft, rDeltas, RedRGB, usedRightColor
            
            Dim iNextOffs As Integer
            If rDeltas.Columns.Count = 1 Then iNextOffs = 2 Else iNextOffs = 1
            Dim rTestNext As Range: Set rTestNext = r.Offset(0, rDeltas.Columns.Count + iNextOffs)
            ColorWithAddressAnchor rTestNext, rightColor
        End If
    End If
End Sub


Private Sub ColorTargetRange(rTargetTopLeft As Range, rDeltas As Range, leftColor As Long, rightColor As Long)
    Dim dMinDelta As Double: dMinDelta = Application.WorksheetFunction.Min(rDeltas)
    Dim dMaxDelta As Double: dMaxDelta = Application.WorksheetFunction.Max(rDeltas)

    If -dMaxDelta < dMinDelta Then
        dMinDelta = -dMaxDelta
    End If
                
    Dim nRows As Integer: nRows = rDeltas.Rows.Count
    Dim nCols As Integer: nCols = rDeltas.Columns.Count
    Dim ixRow As Integer
    Dim ixCol As Integer
    For ixRow = 1 To nRows
        For ixCol = 1 To nCols
            Dim dDelta As Double: dDelta = AsDouble(rDeltas(ixRow, ixCol).Value2)
            Dim col As Long: col = CalcColor(dDelta, dMinDelta, dMaxDelta, leftColor, rightColor)
            Dim rCellToColor As Range: Set rCellToColor = rTargetTopLeft.Offset(ixRow - 1, ixCol - 1)
            rCellToColor.Interior.Color = col
        Next ixCol
    Next ixRow
End Sub


'*************************************
' manual color scale
'*************************************

' http://stackoverflow.com/questions/28217226/creating-a-color-scale-using-vba-avoiding-conditional-formatting

Public Function CalcColor(d As Double, dMin As Double, dMax As Double, colorLow As Long, colorHigh As Long)
    CalcColor = RGB(255, 255, 255)
    If dMin >= 0 Then
        Application.StatusBar = "dMin = " & Format(dMin, "0.000") & " >= 0"
        Exit Function
    End If
    If dMax <= 0 Then
        Application.StatusBar = "dMax = " & Format(dMax, "0.000") & " <= 0"
        Exit Function
    End If
    
    Dim col As Long
    Dim dfrac As Double
    If d >= 0 Then
        dfrac = d / dMax
        col = CalcColorScale(rgbWhite, colorHigh, dfrac)
    Else
        dfrac = d / dMin
        col = CalcColorScale(rgbWhite, colorLow, dfrac)
    End If
    CalcColor = col
End Function

' color1: The starting color as a long
' color2: The end color as a long
' dScale: This is the percentage in decimal of the color.
Public Function CalcColorScale(color1 As Long, color2 As Long, dScale As Double) As Long

    ' Convert the colors to red, green, blue components
    Dim red1 As Long, green1 As Long, blue1 As Long
    red1 = color1 Mod 256
    green1 = (color1 \ 256) Mod 256
    blue1 = (color1 \ 256 \ 256) Mod 256

    Dim red2 As Long, green2 As Long, blue2 As Long
    red2 = color2 Mod 256
    green2 = (color2 \ 256) Mod 256
    blue2 = (color2 \ 256 \ 256) Mod 256

    CalcColorScale = RGB(CalcColorScaleRGB(red1, red2, dScale) _
                        , CalcColorScaleRGB(green1, green2, dScale) _
                        , CalcColorScaleRGB(blue1, blue2, dScale))
End Function

' Calculates the R,G or B for a color between two colors based the percentage between them
' e.g .5 would be halfway between the two colors
 Public Function CalcColorScaleRGB(color1 As Long, color2 As Long, dScale As Double) As Long
    If color2 < color1 Then
        CalcColorScaleRGB = color1 - (Abs(color1 - color2) * dScale)
    ElseIf color2 > color1 Then
        CalcColorScaleRGB = color1 + (Abs(color1 - color2) * dScale)
    Else
        CalcColorScaleRGB = color1
    End If
End Function



'>> Comments

Sub Comment_Worksheet_Calculate()
    Application.ScreenUpdating = False
    Dim rOld As Range: Set rOld = Selection
    Cells.EntireRow.AutoFit
    rOld.Select
    Application.ScreenUpdating = True
End Sub


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


Sub DumpAllCommentsForActiveSheet()
    SetSilentApplicationState
    DumpAllCommentsForSheet ActiveSheet, ActiveCell, False
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
        rAnchor.Cells(nComments, 1).formula = "=GetAddress(" & rCommentCell.Address(External:=True) & ", " & IIf(bExternal, "true", "false") & ")"
        
        ' col 1: show comment text
        rAnchor.Cells(nComments, 2).formula = "=GetComment(" & rAnchor.Cells(nComments, 1).Address & ", true)"
        
    Next
    
    DumpAllCommentsForSheet = nComments
End Function


'*************************************
' helpers
'*************************************

Function GetContentFromActiveCell()
    GetContentFromActiveCell = ActiveCell.Value2
End Function


Sub RemoveAllCommentsFromSheet()
    'MsgBox "deprecated - exiting"
    'Exit Sub
    
    Dim cmt As Comment
    For Each cmt In ActiveSheet.Comments
        cmt.Parent.ClearComments
    Next
End Sub


Sub RemoveNameFromComments(Optional sName As String)
    ' make sure we have a name
    If sName = "" Then sName = "Frank Schuhardt:"
    
    ' standard names in comments have a trailing ":", also remove that
    If Right(sName, 1) <> ":" Then sName = sName & ":"
    
    Dim cmt As Comment
    For Each cmt In ActiveSheet.Comments
        Dim sCommentText As String: sCommentText = cmt.Text
        If StartsWith(sCommentText, sName) Then
            With cmt.Parent
                .ClearComments
                .AddComment Mid(sCommentText, Len(sName) + 2)
            End With
        End If
    Next
End Sub


Sub RemoveFrankSchuhardtFromComments()
    MsgBox "deprecated - exiting"
    Exit Sub
    
    RemoveNameFromComments "Frank Schuhardt:"
End Sub


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


Function GetCommentFromString(sRange As String) As String
    GetCommentFromString = GetComment(Range(sRange))
End Function


'*************************************
' called from AHK
'*************************************

Sub JumpToNextCellWithComment()
    Dim cmt As Comment
    For Each cmt In ActiveSheet.Comments
        Dim rCommentCell As Range: Set rCommentCell = cmt.Parent
        If ThisRangeIsAfterThatRange(rCommentCell, ActiveCell) Then
            rCommentCell.Select
            Exit Sub
        End If
    Next
End Sub


Private Function ThisRangeIsAfterThatRange(rThis As Range, rThat As Range) As Boolean
    If rThis.Row < rThat.Row Then
        ThisRangeIsAfterThatRange = False
    Else
        If rThis.Row > rThat.Row Then
            ThisRangeIsAfterThatRange = True
        Else
            ThisRangeIsAfterThatRange = rThis.Column > rThat.Column
        End If
    End If
End Function


Public Function GetCommentFromActiveCell() As String
    ' from AHK GUI
    On Error Resume Next
    GetCommentFromActiveCell = GetComment(ActiveCell)
End Function


Sub SaveCommentToActiveCell(s)
    ' from AHK GUI
    ActiveCell.ClearComments
    If s <> "" Then ActiveCell.AddComment s
End Sub


'*************************************
' lowlevel set comment
'*************************************

Sub SetComment(r As Range, sNewComment As String, Optional bTouch = True)
    If r.Cells.Count = 1 Then
        r.ClearComments
        r.AddComment sNewComment
        If bTouch Then Touch r
    End If
End Sub


'*************************************
' comment parts
'*************************************

Function GetCommentPartValue(r As Range, vPartKey) As Variant
    Dim sComment As String: sComment = GetComment(r)
    If sComment <> "" Then
        Dim RE As String: RE = FS("^ *(#) *= *(.*)$", vPartKey)
        Dim vPartValue: vPartValue = RegMatch(sComment, RE, 2, False, True)
        Dim bNumeric: bNumeric = IsNumeric(vPartValue) And UCase(vPartKey) <> TODO_ID
        If bNumeric Then GetCommentPartValue = vPartValue + 0 Else GetCommentPartValue = vPartValue
    End If
End Function


Sub ChangeCommentPartValue(r As Range, vPartKey, vNewPartValue)
    Dim sComment As String: sComment = GetComment(r)
    If sComment <> "" Then
        Dim RE As String: RE = FS("^ *(#) *= *(.*)$", vPartKey)
        Dim sFound As String: sFound = RegMatch(sComment, RE, 0, False, True)
        Dim sNewComment As String
        If sFound = "" Then
            If Not EndsWith(sComment, Chr(10)) Then sComment = sComment & Chr(10)
            sNewComment = sComment & vPartKey & "=" & vNewPartValue
        Else
            Dim sPartValue As String: sPartValue = RegMatch(sComment, RE, 1, False, True)
            sNewComment = Replace(sComment, sFound, vPartKey & "=" & vNewPartValue)
        End If
        SetComment r, sNewComment
    End If
End Sub


Sub ReplaceCommentPartKey(r As Range, vOldPartKey, vNewPartKey)
    If r.Cells.Count > 1 Then
        Dim rSingle As Range
        For Each rSingle In r
            ReplaceCommentPartKey rSingle, vOldPartKey, vNewPartKey
        Next
    Else
        Dim sComment As String: sComment = Trim(GetComment(r))
        If sComment <> "" Then
            Dim RE As String: RE = FS("^ *(#) *= *(.*)$", vOldPartKey)
            Dim sFound As String: sFound = RegMatch(sComment, RE, 0, False, True)
            If sFound <> "" Then
                Dim sPartValue As String: sPartValue = RegMatch(sComment, RE, 2, False, True)
                Dim sNewComment As String: sNewComment = Replace(sComment, sFound, vNewPartKey & "=" & sPartValue)
                SetComment r, sNewComment
            End If
        End If
    End If
End Sub




' >> Copying

' ((EPMHNMQ)) adjust EXCEL_LAST_ROW_PJ for pre-10 versions
    
'*************************************
' AdjustTargetRowsAndCopyColumns
'*************************************

Function AdjustTargetRowsAndCopyColumns(sSourceSheet As String, sTargetSheet As String, sHeaderRange, sGuideCol As String, ParamArray asColumns()) As Long
    ' REN for header must be the same for source and target sheet (usually "Header")
    
    SetSilentApplicationState
    On Error GoTo done
    
    ShowStatus SubstituteParams("copying columns from {1} to {2}", sSourceSheet, sTargetSheet)
    
    Dim rSourceHeader As Range: Set rSourceHeader = Worksheets(sSourceSheet).Range(sHeaderRange)
    Dim rTargetHeader As Range: Set rTargetHeader = Worksheets(sTargetSheet).Range(sHeaderRange)
    
    ' ArticleID is complete (e.g. no empty cells), determines # of rows to copy
    Dim nRows As Long: nRows = DetermineColumnRows(rSourceHeader, sGuideCol)
    
    ' expand/contract to new number of rows
    AdjustTargetRows rTargetHeader, nRows
    
    ' new rows now contain the same formulas as the previously last row
    
    ' copy the cells (might override formulas copied above)
    CopyColumns rSourceHeader, rTargetHeader, nRows, asColumns
        
    ' try to set timestamps
    SetTimestamps sTargetSheet
    
    AdjustTargetRowsAndCopyColumns = nRows
done:
    RevertApplicationState
End Function


Sub AdjustTargetRows(rTargetHeader As Range, nNewRows As Long)
    ' expand or contract a range to nNewRows rows
    ' expanding copies cell formulas from previously last row

    Dim rHeader As Range
    Dim nOldRows As Long
    For Each rHeader In rTargetHeader
        If IsEmpty(rHeader) Then Exit For
        Dim sColumn As String: sColumn = rHeader.Value2
        Dim nRowsInColumn As Long: nRowsInColumn = DetermineColumnRows(rTargetHeader, sColumn)
        If nRowsInColumn <> EXCEL_LAST_ROW_PJ Then nOldRows = Max(nOldRows, nRowsInColumn)
    Next
    
    If nOldRows = nNewRows Then
    
        ' nothing to do
        
    Else
    
        If nNewRows < nOldRows Then
        
            ' too many rows - cut back
            Dim nRowsToDelete As Long: nRowsToDelete = nOldRows - nNewRows
            Dim rLastOldToDelete As Range: Set rLastOldToDelete = rTargetHeader.Offset(nOldRows, 0)
            Dim rFirstOldToDelete As Range: Set rFirstOldToDelete = rLastOldToDelete.Offset(-(nRowsToDelete - 1))
        
            rTargetHeader.Worksheet.Range(rFirstOldToDelete, rLastOldToDelete).EntireRow.Delete
        
        Else
        
            Dim nRowsToAdd As Long: nRowsToAdd = nNewRows - nOldRows
            Dim rLastOld As Range: Set rLastOld = rTargetHeader.Offset(nOldRows, 0)
            Dim rFirstNew As Range: Set rFirstNew = rLastOld.Offset(1, 0)
            Dim rLastNew As Range: Set rLastNew = rTargetHeader.Offset(nNewRows, 0)
        
            rLastOld.Copy
            rTargetHeader.Worksheet.Range(rFirstNew, rLastNew).PasteSpecial xlPasteAll
            Application.CutCopyMode = False
        
        End If
    
    End If
End Sub


Sub AdjustTargetRows2(rTargetHeader As Range, nNewRows As Long)
    ' expand or contract a range to nNewRows rows
    ' expanding copies cell formulas from previously last row

    Dim rHeader As Range
    Dim nOldRows As Long
    For Each rHeader In rTargetHeader
        If IsEmpty(rHeader) Then Exit For
        Dim sColumn As String: sColumn = rHeader.Value2
        Dim nRowsInColumn As Long: nRowsInColumn = DetermineColumnRows2(rTargetHeader, sColumn)
        If nRowsInColumn <> -1 Then nOldRows = Max(nOldRows, nRowsInColumn)
    Next
    
    If nOldRows = nNewRows Then
    
        ' nothing to do
        
    Else
    
        If nNewRows < nOldRows Then
        
            ' too many rows - cut back
            Dim nRowsToDelete As Long: nRowsToDelete = nOldRows - nNewRows
            Dim rLastOldToDelete As Range: Set rLastOldToDelete = rTargetHeader.Offset(nOldRows, 0)
            Dim rFirstOldToDelete As Range: Set rFirstOldToDelete = rLastOldToDelete.Offset(-(nRowsToDelete - 1))
        
            rTargetHeader.Worksheet.Range(rFirstOldToDelete, rLastOldToDelete).EntireRow.Delete
        
        Else
        
            Dim nRowsToAdd As Long: nRowsToAdd = nNewRows - nOldRows
            Dim rLastOld As Range: Set rLastOld = rTargetHeader.Offset(nOldRows, 0)
            Dim rFirstNew As Range: Set rFirstNew = rLastOld.Offset(1, 0)
            Dim rLastNew As Range: Set rLastNew = rTargetHeader.Offset(nNewRows, 0)
        
            rLastOld.Copy
            rTargetHeader.Worksheet.Range(rFirstNew, rLastNew).PasteSpecial xlPasteAll
            Application.CutCopyMode = False
        
        End If
    
    End If
End Sub


Sub CopyColumn(rSourceHeader As Range, sSourceColumn As String, rTargetHeader As Range, sTargetColumn As String, nRows As Long)
    ' copy nRows cells in a named source column to a target range
    ' column names in source and target ranges don't have to be the same

    Dim ixSourceColumn As Long: ixSourceColumn = match(sSourceColumn, rSourceHeader)
    Debug.Assert ixSourceColumn > 0
    Dim ixTargetColumn As Long: ixTargetColumn = match(sTargetColumn, rTargetHeader)
    Debug.Assert ixTargetColumn > 0
    
    Dim rSourceTop As Range: Set rSourceTop = rSourceHeader.Cells(ixSourceColumn).Offset(1, 0)
    Dim rTargetTop As Range: Set rTargetTop = rTargetHeader.Cells(ixTargetColumn).Offset(1, 0)
        
    Dim rSourceBottom As Range: Set rSourceBottom = rSourceTop.Offset(nRows - 1, 0)
    Dim rTargetBottom As Range: Set rTargetBottom = rTargetTop.Offset(nRows - 1, 0)
    
    Dim rSource As Range: Set rSource = rSourceTop.Worksheet.Range(rSourceTop, rSourceBottom)
    Dim rTarget As Range: Set rTarget = rTargetTop.Worksheet.Range(rTargetTop, rTargetBottom)
    
    rTarget.Value2 = rSource.Value2
End Sub


Sub CopyColumns(rSourceHeader As Range, rTargetHeader As Range, nRows As Long, ParamArray pasColumns())
    
    ' for AdjustTargetRowsAndCopyColumns
    Dim asColumns(): asColumns = ParamArrayDelegated(pasColumns)
    
    Dim nPassedColumns As Long: nPassedColumns = UBound(asColumns) - LBound(asColumns) + 1
    Debug.Assert nPassedColumns Mod 2 = 0
    Dim ixColumn As Long
    For ixColumn = LBound(asColumns) To UBound(asColumns) Step 2
        Dim sSourceColumn As String: sSourceColumn = asColumns(ixColumn)
        Dim sTargetColumn As String: sTargetColumn = asColumns(ixColumn + 1)
        CopyColumn rSourceHeader, sSourceColumn, rTargetHeader, sTargetColumn, nRows
    Next
End Sub



'>> Dates

Function AsDate(v) As Double
    If IsDate(v) Then
        ' Nur dieser Umweg funktioniert, sonst wird ein Double zurÃ¼ckgegeben, der dem Datum entspricht - - nur ohne "."
        AsDate = DateSerial(Year(v), Month(v), Day(v)) + TimeSerial(Hour(v), Minute(v), Second(v))
    Else
        If IsNumeric(v) Then
            AsDate = v
        Else
            ' funktioniert das?
            Dim s As String: s = v
            AsDate = s
        End If
    End If
End Function


Function SafeDate(v As Variant) As Variant
    Const RE1 = "([0-9][0-9])([0-9][0-9])([0-9][0-9][0-9][0-9])"
    On Error Resume Next
    Dim d As Double: d = AsDate(v)
    If d = 0 Then
        SafeDate = "?"
    Else
        If d < 50000 Then
            SafeDate = d
        Else
            ' assume format DDMMYYYY
            If RegMatch(v, RE1) Then
                SafeDate = DateSerial(RegMatch(v, RE1, 3), RegMatch(v, RE1, 2), RegMatch(v, RE1, 1))
            End If
        End If
    End If
End Function


'*************************************
' days
'*************************************

Function WeekdayGerman(dt As Double) As String
    Dim wd As Integer: wd = Weekday(dt, vbMonday)
    Select Case wd
        Case 1
            WeekdayGerman = "Montag"
        Case 2
            WeekdayGerman = "Dienstag"
        Case 3
            WeekdayGerman = "Mittwoch"
        Case 4
            WeekdayGerman = "Donnerstag"
        Case 5
            WeekdayGerman = "Freitag"
        Case 6
            WeekdayGerman = "Samstag"
        Case 7
            WeekdayGerman = "Sonntag"
    End Select
End Function


Function WeekdayEnglish(dt As Double) As String
    Dim wd As Integer: wd = Weekday(dt, vbMonday)
    Select Case wd
        Case 1
            WeekdayEnglish = "Monday"
        Case 2
            WeekdayEnglish = "Tuesday"
        Case 3
            WeekdayEnglish = "Wednesday"
        Case 4
            WeekdayEnglish = "Thursday"
        Case 5
            WeekdayEnglish = "Friday"
        Case 6
            WeekdayEnglish = "Saturday"
        Case 7
            WeekdayEnglish = "Sunday"
    End Select
End Function


'*************************************
' weeks
'*************************************

Sub ParseWeek(sWeekInYear As String, ByRef iYear As Integer, ByRef iWeek As Integer)
    Const RE_Parse = "^(20[0-9][0-9])-([0-9]+)$"
    iYear = RegMatch(sWeekInYear, RE_Parse, 1)
    iWeek = RegMatch(sWeekInYear, RE_Parse, 2)
End Sub


Function WeeksCount(sWeekStringFrom As String, sWeekStringTo As String) As Integer
    WeeksCount = (WeekStringToDate(sWeekStringTo) - WeekStringToDate(sWeekStringFrom)) / 7 + 1
End Function


Function DateToWeekNumber(dt As Double) As Integer
    DateToWeekNumber = DatePart("ww", dt, vbMonday, vbFirstFourDays)
End Function


Function DateToWeekString(dt As Double) As String
    ' e.g. 2016-05
    
    Dim iWeek As Integer: iWeek = DateToWeekNumber(dt)
    Dim iYear As Integer: iYear = Year(dt)
    If iWeek = 1 And Month(dt) = 12 Then iYear = iYear + 1
    If iWeek = 52 And Month(dt) = 1 Then iYear = iYear - 1
    DateToWeekString = iYear & "-" & Format(iWeek, "00")
End Function


Function YearWeekToWeekString(iYear As Integer, iWeek As Integer) As String
    YearWeekToWeekString = iYear & "-" & Format(iWeek, "00")
End Function


Function WeekToDate(iYear As Integer, iWeek As Integer) As Double
    WeekToDate = DateSerial(iYear, 1, -2) - Weekday(DateSerial(iYear, 1, 3)) + iWeek * 7
End Function


Function WeekStringToDate(sWeekString As String) As Double
    Dim iYear As Integer: iYear = Left(sWeekString, 4) + 0
    Dim iWeek As Integer: iWeek = Right(sWeekString, 2) + 0
    WeekStringToDate = WeekToDate(iYear, iWeek)
End Function


Function WeekStringPlusWeeks(sWeekString As String, nWeeks As Integer) As String
    WeekStringPlusWeeks = DateToWeekString(WeekStringToDate(sWeekString) + nWeeks * 7)
End Function


'*************************************
' overlap weeks <-> month
'*************************************

Function FirstInMonth(dt As Double) As Double
    FirstInMonth = DateSerial(Year(dt), Month(dt), 1)
End Function

Function LastInMonth(dt As Double) As Double
    LastInMonth = DateSerial(Year(dt), Month(dt) + 1, 1) - 1
End Function

Function DaysInMonth(dt As Double) As Double
    DaysInMonth = LastInMonth(dt) - FirstInMonth(dt) + 1
End Function


'*************************************
' overlap weeks <-> month
'*************************************

Function DaysOfWeekInMonths(rRefDates As Range, iWeek As Integer) As Variant
    Dim adtRefDates(): adtRefDates = RangeToVector(rRefDates)
        
    Dim aDays(): ReDim aDays(1 To ArrayCount(adtRefDates))
    Dim n As Integer
    Dim ix As Integer
    For ix = LBound(adtRefDates) To UBound(adtRefDates)
        Dim dtRefDate As Double: dtRefDate = adtRefDates(ix)
        Dim iDaysOfWeekInMonth As Integer: iDaysOfWeekInMonth = DaysOfWeekInMonth(dtRefDate, iWeek)
        n = n + 1
        aDays(n) = iDaysOfWeekInMonth
    Next
    
    DaysOfWeekInMonths = aDays
End Function


Function DaysOfWeekInMonth(dtRefDate As Double, iWeek As Integer) As Integer
    Dim iYear As Integer: iYear = Year(dtRefDate)
    Dim iMonth As Integer: iMonth = Month(dtRefDate)
    
    Dim iYearOfStartWeek As Integer
    If iMonth = 1 And iWeek > 50 Then iYearOfStartWeek = iYear - 1 Else iYearOfStartWeek = iYear
    
    Dim dtFirstInWeek As Double: dtFirstInWeek = WeekToDate(iYearOfStartWeek, iWeek)
    Dim dtLastInWeek As Double: dtLastInWeek = dtFirstInWeek + 6
    Dim dtFirstInMonth As Double: dtFirstInMonth = DateSerial(iYear, iMonth, 1)
    Dim dtLastInMonth As Double: dtLastInMonth = DateSerial(iYear, iMonth + 1, 1) - 1
    
    If dtLastInWeek < dtFirstInMonth Or dtFirstInWeek > dtLastInMonth Then
        DaysOfWeekInMonth = 0
    Else
        
        If dtFirstInWeek < dtFirstInMonth Then
        
            DaysOfWeekInMonth = 7 - (dtFirstInMonth - dtFirstInWeek)
        
        Else
        
            If dtLastInWeek > dtLastInMonth Then
            
                DaysOfWeekInMonth = 7 - (dtLastInWeek - dtLastInMonth)
                
            Else
            
                DaysOfWeekInMonth = 7
            End If
        End If
    End If
End Function



' >> Dents

'*************************************
' sheet analysis
'*************************************

' called by Autohotkey
Sub ShowSheetReferences()
    Dim cRefs As Collection: Set cRefs = CollectSheetReferences(Cells)
    Dim sRefs As String: sRefs = CollectionToString(cRefs, Chr(13) & Chr(10))
    MsgBox sRefs
    PutOnClipboard sRefs
End Sub


' this does not use Dents but is significantly faster
Function CollectSheetReferences(rToSearchIn As Range) As Collection
    Dim dRefs As New Dictionary
    With rToSearchIn
        Dim r As Range: Set r = .Find("!", LookIn:=xlFormulas)
        If Not r Is Nothing Then
            Dim firstAddress As String: firstAddress = r.Address
            Do
                Dim sFormula As String: sFormula = r.formula
                
                ' find ref
                Do
                
                    Dim sRef As String: sRef = RegMatch(sFormula, "('[^']+')!", 1)
                
                    ' if not found a '...' reference then try to find a ref w/o the '
                    If sRef = "" Then sRef = RegMatch(sFormula, "([^(=!]+)!", 1)
                    
                    ' add the ref if not already found earlier
                    If sRef <> "" Then
                    
                        'If Contains(sRef, "#") Then Stop
                        
                        If Not dRefs.Exists(sRef) Then dRefs.Add sRef, sRef
                        
                        sFormula = Replace(sFormula, sRef & "!", "")
                    End If
                Loop Until sRef = ""
                
                Set r = .FindNext(r)
            Loop While Not r Is Nothing And r.Address <> firstAddress
        End If
    End With
    
    Set CollectSheetReferences = DictionaryKeysToCollection(dRefs)
End Function


' better functionality if we need to have all cells which contain a link
' basically like CollectSheetReferences()
Function CollectSheetReferencesDictionary(rToSearchIn As Range) As Dictionary
    Dim dRefs As New Dictionary
    With rToSearchIn
        Dim r As Range: Set r = .Find("!", LookIn:=xlFormulas)
        If Not r Is Nothing Then
            Dim firstAddress As String: firstAddress = r.Address
            Do
                Dim sFormula As String: sFormula = r.formula
                
                ' find ref
                Do
                
                    Dim sRef As String: sRef = RegMatch(sFormula, "('[^']+')!", 1)
                
                    ' if not found a '...' reference then try to find a ref w/o the '
                    If sRef = "" Then sRef = RegMatch(sFormula, "([^(=!]+)!", 1)
                    
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
                        cRefs.Add r
                        
                        sFormula = Replace(sFormula, sRef & "!", "")
                    End If
                Loop Until sRef = ""
                
                Set r = .FindNext(r)
            Loop While Not r Is Nothing And r.Address <> firstAddress
        End If
    End With
    
    Set CollectSheetReferencesDictionary = dRefs
End Function


' do not use this - - too slow for big sheets (i.e. many formulas)
' use ShowSheetReferences() instead
Sub ShowSheetReferencesWithDents()
    Application.ScreenUpdating = False
    Dim dRefs As New Dictionary
    Dim rFormulas As Range: Set rFormulas = ActiveSheet.Cells.SpecialCells(xlCellTypeFormulas)
    Dim rFormula As Range
    For Each rFormula In rFormulas
        
        Dim cDents As Collection
        
        Set cDents = GetDents(True, rFormula, 0)
        
        Dim vPrecedent
        For Each vPrecedent In cDents
            Dim sRef As String: sRef = RegMatch(AsString(vPrecedent), "('?[^'!]+'?)!", 1)
            If Not dRefs.Exists(sRef) Then
            
                dRefs.Add sRef, sRef
            
            End If
        Next
        
    Next
    
    Dim cRefs As Collection: Set cRefs = DictionaryKeysToCollection(dRefs)
    MsgBox CollectionToString(cRefs, Chr(13) & Chr(10))
    
    Application.ScreenUpdating = True
End Sub


Sub CopySheetReferencesToClipboard()
    Application.ScreenUpdating = False
    
    Dim rCurrent As Range: Set rCurrent = ActiveCell
    
    On Error Resume Next
    Dim rFormulas As Range: Set rFormulas = ActiveSheet.Range("A1").SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If rFormulas Is Nothing Then Exit Sub
    
    Dim dict As New Dictionary
    
    Dim rToCheck As Range
    For Each rToCheck In rFormulas
        Dim c As Collection: Set c = GetDents(True, rToCheck)
        Dim vAddress
        For Each vAddress In c
            Dim sSheet As String: sSheet = RegMatch(vAddress, "'?\[[A-Za-z0-9. -]+?\]([^!']+?)'?!", 1)
            If Not dict.Exists(sSheet) Then dict.Add sSheet, sSheet
        Next
    Next
    
    Dim v: Dim cSheets As New Collection
    For Each v In dict.Keys
        cSheets.Add v
    Next
    Dim s As String: s = CollectionToString(cSheets, Chr(13) & Chr(10))
    
    PutInClipboard s
    
    Application.ScreenUpdating = True
    rCurrent.Select
End Sub


'*************************************
' Standard Ctrl-D, Ctrl-P
'*************************************

Public Sub DoAllPrecedents()
    DoAllDents True, DENTS_TIMEOUT_SECONDS
End Sub

Public Sub DoAllDependents()
    DoAllDents False, DENTS_TIMEOUT_SECONDS
End Sub


'won't navigate through precedents in closed workbooks
'won't navigate through precedents in protected worksheets
'won't identify precedents on hidden sheets
Public Sub DoAllDents(bPrecedents As Boolean, Optional iTimeoutSeconds As Integer)
    Application.ScreenUpdating = False
    Dim cAllDents As Collection: Set cAllDents = GetDents(bPrecedents, ActiveCell, iTimeoutSeconds)
    Application.ScreenUpdating = True
    
    Application.StatusBar = False
    
    Dim sDent As String: If bPrecedents Then sDent = "precedent" Else sDent = "dependent"
    Dim bNoDents As Boolean: bNoDents = cAllDents Is Nothing: If Not bNoDents Then bNoDents = cAllDents.Count = 0
    If bNoDents Then
        Application.StatusBar = TrimCellAddress(ActiveCell.Address(External:=True)) & " has no " & sDent & " cells."
    Else
        Dim uf As New UserFormDents
        uf.Caption = sDent & "s"
                
        Dim v
        For Each v In Sorted(Unique(cAllDents))
            uf.ListBox1.AddItem TrimCellAddress(v)
        Next
        
        ' set to first element
        uf.ListBox1.ListIndex = 0
        
        uf.StartUpPosition = 0
        uf.Top = 50
        
        If g_UserFormDentsLeft = 0 Then
            uf.Left = USERFORMDENTS_LEFT
        Else
            uf.Left = g_UserFormDentsLeft
        End If
        
        uf.Show
    End If

    Application.ScreenUpdating = True
End Sub

  
Function GetDents(bPrecedents As Boolean, ByRef rToCheck As Range, iTimeoutSeconds As Integer) As Collection
 
    ' get either precedents or descendents
 
    Dim cAllDents As New Collection
 
    Dim rngCell As Range
    Dim rngFormulas As Range
 
    Dim dStartTime As Double: dStartTime = Now()
 
    If Not rToCheck.Worksheet.ProtectContents Then
        If rToCheck.Cells.Count > 1 Then
            On Error Resume Next
            Set rngFormulas = rToCheck.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
        Else
            ' was: If rToCheck.HasFormula Then Set rngFormulas = rToCheck
            Set rngFormulas = rToCheck
        End If

        ' was: only start iterating if rngFormulas is not nothing
        Debug.Assert Not (rngFormulas Is Nothing)
 
        For Each rngCell In rngFormulas.Cells
            GetCellDents bPrecedents, rngCell, cAllDents, dStartTime, iTimeoutSeconds
        Next rngCell
        On Error GoTo err
        rngFormulas.Worksheet.ClearArrows
    End If
 
    Set GetDents = cAllDents
    Exit Function
err:
    Set GetDents = Nothing
End Function
 
 
Private Sub GetCellDents(bPrecedents As Boolean, ByRef rCell As Range, ByRef cAllDents As Collection, ByRef dStartTime As Double, iTimeoutSeconds As Integer)
 
    Dim lngArrow As Long
    Dim lngLink As Long
    Dim blnNewArrow As Boolean
    Dim sDentAddress As String
    Dim rDent As Range
 
    If bPrecedents Then rCell.ShowPrecedents Else rCell.ShowDependents
    Do
        lngArrow = lngArrow + 1
        blnNewArrow = True
        lngLink = 0
 
        Do
        
            If iTimeoutSeconds > 0 Then
                Dim dElapsedSeconds As Double: dElapsedSeconds = (Now() - dStartTime) * 86400
                If dElapsedSeconds > iTimeoutSeconds Then
                    Dim res: res = MsgBox("timeout - continue?", vbExclamation + vbYesNo)
                    If res = vbNo Then
                        Exit Sub
                    End If
                    dStartTime = Now()
                End If
            End If
        
            lngLink = lngLink + 1
  
            On Error Resume Next
            Set rDent = rCell.NavigateArrow(bPrecedents, lngArrow, lngLink)
 
            If err.Number <> 0 Then
                Exit Do
            End If
 
            On Error GoTo 0
            
            Dim sCellAddress As String: sCellAddress = rCell.Address(False, False, xlA1, True)
            sDentAddress = rDent.Address(False, False, xlA1, True)
            
            ' check if we have a self-link
            ' NOTE: if we check from, say, A1 and this is part of a fused range, then the sDentAddress will be a range
            Dim bSelfLink As Boolean: bSelfLink = Left(sDentAddress & ":", Len(sCellAddress) + 1) = sCellAddress & ":"
            If bSelfLink Then
                ' (1) see below
                Exit Do
            Else
 
                blnNewArrow = False
 
                'If Not cAllDents.Exists(sDentAddress) Then
                    cAllDents.Add sDentAddress
                'End If
            End If
        Loop
 
        If blnNewArrow Then Exit Do
    Loop
End Sub


Function DentsToString(bPrecendent As Boolean, rToCheck As Range, Optional delim As String, Optional iTimeoutSeconds As Integer) As String
    ' funktioniert vermutlich nicht, weil (1) oben mÃ¶glicherweise nicht fÃ¼r functions verwendet werden kann
    Dim cDents As Collection: Set cDents = GetDents(bPrecendent, rToCheck, iTimeoutSeconds)
    DentsToString = CollectionToString(cDents, delim)
End Function


Function PrecedentsToString(rToCheck As Range, Optional delim As String, Optional iTimeoutSeconds As Integer) As String
    PrecedentsToString = DentsToString(True, rToCheck, delim, iTimeoutSeconds)
End Function


Public Function DependentsToString(rToCheck As Range, Optional delim As String, Optional iTimeoutSeconds As Integer) As String
    DependentsToString = DentsToString(False, rToCheck, delim, iTimeoutSeconds)
End Function


Sub FillDents(bPrecedent As Boolean, rMultipleCellsToCheck As Range, rToFill As Range)
    Debug.Assert rMultipleCellsToCheck.Rows.Count = rToFill.Rows.Count
    Debug.Assert rMultipleCellsToCheck.Columns.Count = rToFill.Columns.Count
    Dim ixRow As Integer
    For ixRow = 1 To rMultipleCellsToCheck.Rows.Count
        Dim ixCol As Integer
        For ixCol = 1 To rMultipleCellsToCheck.Columns.Count
            Dim sDents As String: sDents = DentsToString(bPrecedent, rMultipleCellsToCheck.Cells(ixRow, ixCol))
            rToFill.Cells(ixRow, ixCol) = sDents
        Next
    Next
End Sub


Sub FillDependents(rMultipleCellsToCheck As Range, rToFill As Range)
    FillDents False, rMultipleCellsToCheck, rToFill
End Sub


Sub FillDependentsTest()
    Debug.Assert Selection.Areas.Count = 2
    Dim rMultipleCellsToCheck As Range: Set rMultipleCellsToCheck = Selection.Areas(1)
    Dim rToFill As Range: Set rToFill = Selection.Areas(2)
    FillDependents rMultipleCellsToCheck, rToFill
End Sub



'>> Files

'*************************************
' tests
'*************************************

Sub SplitTest()
    Dim sPath As String: sPath = "s:\Dataroom\"
    Dim vFilenames: vFilenames = SortVector(LoadFileToArray("s:\dataroom\alle.txt"))
    Dim av(): av = SplitQualifiedFilenames(vFilenames, Len(sPath), True)
End Sub


'*************************************
' special folders
'*************************************

Function GetTempFolder() As String
    ' https://www.rondebruin.nl/win/s3/win027.htm
    Dim FSO As Object, TmpFolder As Object
    Set FSO = CreateObject("scripting.filesystemobject")
    Set TmpFolder = FSO.GetSpecialFolder(2)
    Dim sfn As String: sfn = TmpFolder
    GetTempFolder = EnsureTrailingBackslash(sfn)
End Function


'*************************************
' physical translation of SUBST
'*************************************

Function GetPhysicalPath(sPath As String) As String
    Dim ixColon As Integer: ixColon = InStr(1, sPath, ":")
    If ixColon = 0 Then
        GetPhysicalPath = sPath
    Else
        Dim sDrive As String: sDrive = UCase(Left(sPath, ixColon))
        If Not SafeSubsts.Exists(sDrive) Then
            GetPhysicalPath = sPath
        Else
        
            Dim sSubstPath As String: sSubstPath = SafeSubsts(sDrive)
            GetPhysicalPath = sSubstPath & Mid(sPath, ixColon + 1)
        End If
    End If
End Function
 
 
Function SafeSubsts() As Dictionary
    If g_Substs Is Nothing Then
        Set g_Substs = New Dictionary
        FillSubstsDictionary
    End If
    Set SafeSubsts = g_Substs
End Function


Private Sub FillSubstsDictionary()
    Const PATTERN = "^([A-Z]:)\\: => (.+)"
    
    Dim cSubst As Collection: Set cSubst = ShellRun("cmd.exe /c subst")
    Dim sSubst
    For Each sSubst In cSubst
        If sSubst = "" Then Exit Sub
        Dim sDrive As String: sDrive = RegMatch(sSubst, PATTERN, 1)
        Dim sSubstPath As String: sSubstPath = RegMatch(sSubst, PATTERN, 2)
        If Not SafeSubsts.Exists(sDrive) Then SafeSubsts.Add sDrive, sSubstPath
    Next
End Sub


Public Function ShellOpen(sfn As String)
    CreateObject("Shell.Application").Open (sfn)
End Function


Public Function ShellRun(sCmd As String, Optional bWaitOnReturn = True, Optional bQuoteCmd As Boolean = True) As Collection
    Const TEMPFILE = "$$$"
    
    Dim oShell: Set oShell = CreateObject("WScript.Shell")
    
    If bQuoteCmd Then sCmd = Quoted(sCmd)
    
    If Not bWaitOnReturn Then
        oShell.Run sCmd, 0, False
    Else
    
        ' run SUBST, don't show window, redirect to temporary file
        ' https://ss64.com/vb/run.html
        ' 05.02.18 seltsam, scheint gequoted nicht zu funktionieren
        'Dim res: res = oShell.Run("""" & sCmd & """" & " > " & TEMPFILE, 0, True)
        Dim res: res = oShell.Run(sCmd, 0, True)
        
        If Not FileExists(TEMPFILE) Then
        
            ' design error
            'Debug.Assert False
            
            ' 28.4.17 removed the assertion - - doesn't work in VDI
            
        Else
            Open TEMPFILE For Input As #1
            
            Set ShellRun = New Collection
            
            Do Until EOF(1)
                Dim textline: Line Input #1, textline
                ShellRun.Add textline
            Loop
            Close #1
            
            DeleteFile TEMPFILE
        End If
    End If
End Function


'*************************************
' file search
'*************************************

Function WindowsSearchInScope(sSearch As String, sScopePath As String) As Collection
    ' http://www.online-excel.de/excel/singsel_vba.php?f=135
    
    Dim oAdoConnection As New ADODB.Connection
    Dim oAdoRecordset As New ADODB.Recordset
    Dim sAdoConnectString As String
    Dim sQuery As String
    
    On Error GoTo err
    sAdoConnectString = "provider=search.collatordso;extended properties=â€™application=windowsâ€™;"
    oAdoConnection.Open sAdoConnectString
    
    sScopePath = GetPhysicalPath(sScopePath)
    
    sQuery = "SELECT System.ItemName, System.DateCreated FROM SYSTEMINDEX WHERE SCOPE = '" & sScopePath & "' AND CONTAINS('" & sSearch & "')"
    With oAdoRecordset
        .source = sQuery
        .ActiveConnection = oAdoConnection
        .Open
    End With
    
    Set WindowsSearchInScope = New Collection
    
    Do Until oAdoRecordset.EOF
        Dim value: value = oAdoRecordset.Fields(0).value
        WindowsSearchInScope.Add value
        oAdoRecordset.MoveNext
    Loop
    
done:
    On Error Resume Next ' Sehr Faul
    oAdoRecordset.Close
    oAdoConnection.Close
    Set oAdoRecordset = Nothing
    Set oAdoConnection = Nothing
    Exit Function
err:
    'MsgBox "Fehler: " & err.Description
    Resume done
End Function


'*************************************
' backslash handling
'*************************************

Sub SafeAddBackslash(ByRef sPath As String)
    sPath = Trim(sPath)
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
End Sub

Function EnsureTrailingBackslash(sfn As String) As String
    SafeAddBackslash sfn
    EnsureTrailingBackslash = sfn
End Function


'*************************************
' get file parts
'*************************************

' ACHTUNG: funktioniert nicht alles wie erwartet, wg. Scripting.FileSystemObject kann offensichtlich nicht richtig extensions abtrennen bzw.
'    erkennt/akzeptiert Punkte nicht als Teile von Pfaden

Function GetFilename(sfn As String) As String
    Dim x: x = Split(sfn, Application.PathSeparator)
    GetFilename = x(UBound(x))
End Function

Function GetFilepath(sfn As String) As String
    Dim sFilename As String: sFilename = GetFilename(sfn)
    GetFilepath = Left(sfn, Len(sfn) - Len(sFilename))
End Function

Function GetFileExtension(sfn As String) As String
    Dim FSO As New Scripting.FileSystemObject
    GetFileExtension = FSO.GetExtensionName(sfn)
End Function

Function GetFileBasename(sfn As String) As String
    Dim FSO As New Scripting.FileSystemObject
    GetFileBasename = FSO.GetBaseName(sfn)
End Function

Function GetAbsolutePathName(sfn As String) As String
    Dim FSO As New Scripting.FileSystemObject
    GetAbsolutePathName = FSO.GetAbsolutePathName(sfn)
End Function


'*************************************
' split filename(s) into parts
'*************************************

' used in Vertragsdatenbank

Public Function SplitQualifiedFilenames(vFilenames, Optional nIgnoreLeadingChars As Integer, Optional bIgnoreFolder As Boolean) As Variant()
    Dim aFilenames()
    If TypeOf vFilenames Is Range Then
        Dim rFilenames As Range: Set rFilenames = vFilenames
        Debug.Assert rFilenames.Columns.Count = 1
        aFilenames = RangeToVector(rFilenames)
    Else
        aFilenames = vFilenames
        Debug.Assert IsVector(aFilenames, True)
    End If
        
    Dim nFilenameRows As Long: nFilenameRows = ArrayCount(aFilenames, 1)
    
    Dim nCols As Integer
    On Error Resume Next
    nCols = Application.Caller.Columns.Count
    On Error GoTo 0
    If nCols = 0 Then nCols = 8
    
    Dim nUsedFilenameRows As Long
    
    Dim sfnOriginal As String
    Dim iFileType As Integer
    Dim bIsFolder As Boolean
    
    Dim ixFilename As Long
    Dim vFilename
    For ixFilename = 1 To nFilenameRows
        vFilename = aFilenames(ixFilename)
    
        If IsError(vFilename) Then
        
            ' ignore
            
        Else
    
            sfnOriginal = vFilename
            If sfnOriginal = "" Then
            
                ' ignore
            
            Else
            
                iFileType = GetFileType(sfnOriginal)
                
                bIsFolder = iFileType = 0
                If bIsFolder And bIgnoreFolder Then
                
                    ' ignore
                    
                Else
                
                    nUsedFilenameRows = nUsedFilenameRows + 1
                End If
            End If
        End If
    Next ixFilename
        
    Dim aResult()
    ReDim aResult(1 To nUsedFilenameRows, 1 To nCols)
    
    Dim nMaxParts As Integer
        
    Dim ixRow As Long
    For ixFilename = 1 To nFilenameRows
        vFilename = aFilenames(ixFilename)
    
        If IsError(vFilename) Then
        
            ' ignore
            
        Else
    
            sfnOriginal = vFilename
            If sfnOriginal = "" Then
            
                ' ignore
            
            Else
            
                iFileType = GetFileType(sfnOriginal)
                
                bIsFolder = iFileType = 0
                If bIsFolder And bIgnoreFolder Then
                
                    ' ignore
                    
                Else
                
                    Dim sfn As String: sfn = sfnOriginal
                    If nIgnoreLeadingChars > 0 Then sfn = Mid(sfn, nIgnoreLeadingChars + 1)
                    
                    Dim aParts() As String: aParts = Split(sfn, "\")
                    
                    ' result from Split() is 0-based
                    Debug.Assert LBound(aParts) = 0
                    Dim nParts As Integer: nParts = UBound(aParts) + 1
                    If nParts > nMaxParts Then nMaxParts = nParts
                    
                    ixRow = ixRow + 1
                    
                    aResult(ixRow, 1) = sfnOriginal
                    aResult(ixRow, 2) = iFileType
                    aResult(ixRow, 3) = GetFilename(sfn)
                    
                    Dim ixPart As Integer
                    Dim nMinCols As Integer: nMinCols = Min(nCols - 3, nParts)
                    For ixPart = 1 To nCols - 3
                        Dim bHasMoreParts As Boolean: bHasMoreParts = ixPart <= nMinCols
                        If bHasMoreParts Then
                            aResult(ixRow, ixPart + 3) = aParts(ixPart - 1)
                        Else
                            aResult(ixRow, ixPart + 3) = ""
                        End If
                    Next
                    
                End If
            End If
        End If
    Next
    
    GoTo skipPruning
    
    Debug.Assert nUsedFilenameRows = ixRow
    ' Dim nUsedFilenameRows As long: nUsedFilenameRows = ixRow
    
    Dim nRows As Integer: nRows = 0
    On Error Resume Next
    nRows = Min(Application.Caller.Rows.Count, nUsedFilenameRows)
    On Error GoTo 0
    If nRows = 0 Then nRows = nUsedFilenameRows
    
    While ixRow < nRows
        ixRow = ixRow + 1
        For ixPart = 1 To nCols
            aResult(ixRow, ixPart) = ""
        Next
    Wend
    
skipPruning:
    
    If nMaxParts + 3 > nCols Then MsgBox "too few columns (" & nCols & "), needs more (" & nMaxParts + 3 & ")"
    
    SplitQualifiedFilenames = aResult
End Function


'*************************************
' file / path exists
'*************************************

Function FileExists(sfn As String) As Boolean
   ' http://stackoverflow.com/questions/67835/deleting-a-file-in-vba
    'FileExists = dir(sfn) <> ""
    
    ' 22.04.21 schneller, allerding nur mit prefixing w/ "\\?\" in der Lage, long filenames korrekt zu verarbeiten
    ' https://stackoverflow.com/questions/38432250/access-files-with-long-paths-over-260
    
    Dim FSO As New FileSystemObject
    FileExists = FSO.FileExists(sfn)
End Function


Function DirectoryExists(sfnPath As String) As Boolean
    'DirectoryExists = dir(sfnPath, vbDirectory) <> ""
    ' 22.04.21 assumption: ebenfalls schneller
    Dim FSO As New FileSystemObject
    DirectoryExists = FSO.FolderExists(sfnPath)
End Function

Function PathExists(sfnPath As String) As Boolean
    PathExists = DirectoryExists(sfnPath)
End Function


'*************************************
' delete / copy / rename
'*************************************

Sub DeleteFile(ByVal sfnToDelete As String)
    ' http://stackoverflow.com/questions/67835/deleting-a-file-in-vba
    If FileExists(sfnToDelete) Then 'See above
        SetAttr sfnToDelete, vbNormal
        Kill sfnToDelete
    End If
End Sub

Sub SafeFileCopy(sfnSource As String, sfnTarget As String, Optional bErrorOnAlreadyExists As Boolean)
    If sfnSource = "" And sfnTarget = "" Then Exit Sub
    Debug.Assert FileExists(sfnSource)
    If bErrorOnAlreadyExists Then Debug.Assert bErrorOnAlreadyExists And FileExists(sfnTarget)
    FileCopy sfnSource, sfnTarget
End Sub

Sub SafeFileRename(sfnOld As String, sfnNew As String, Optional bErrorOnAlreadyExists As Boolean)
    Debug.Assert FileExists(sfnOld)
    If FileExists(sfnNew) Then
        Debug.Assert Not bErrorOnAlreadyExists
    Else
        Name sfnOld As sfnNew
    End If
End Sub

Sub ForceFileRename(sfnOld As String, sfnNew As String)
    Name sfnOld As sfnNew
End Sub


'*************************************
' Macros delete / copy / rename
'*************************************

Sub DoDelete()
    Dim r As Range: Set r = Selection
    Debug.Assert r.Columns.Count = 1
    Dim v
    For Each v In r.Cells
        Dim sfnToDelete As String: sfnToDelete = v
        DeleteFile sfnToDelete
    Next
End Sub


Sub DoCopy()
    Dim r As Range: Set r = Selection
    Debug.Assert r.Columns.Count = 2
    Dim rRow As Range
    For Each rRow In r.Rows
        Dim sfnSource As String: sfnSource = rRow.Cells(1)
        Dim sfnTarget As String: sfnTarget = rRow.Cells(2)
        SafeFileCopy sfnSource, sfnTarget
    Next
End Sub


Sub DoRename()
    Dim r As Range: Set r = Selection
    Debug.Assert r.Columns.Count = 2
    Dim rRow As Range
    For Each rRow In r.Rows
        Dim sfnOld As String: sfnOld = rRow.Cells(1)
        Dim sfnNew As String: sfnNew = rRow.Cells(2)
        If FileExists(sfnOld) Then
            SafeFileRename sfnOld, sfnNew
        End If
    Next
End Sub


'*************************************
' "standard" filetype
'*************************************

' u.a. fÃ¼r Vertragsdatenbank coloring

Function GetFileType(sfn As String) As Integer
    Dim FSO As New Scripting.FileSystemObject
    Dim bIsFolder As Boolean: bIsFolder = FSO.FolderExists(sfn)
    If bIsFolder Then
        GetFileType = 0
    Else
        'Dim sBasename As String: sBasename = fso.GetBaseName(sfn)
        Dim sExtension As String: sExtension = LCase(FSO.GetExtensionName(sfn))
        Select Case sExtension
            Case "pdf"
                GetFileType = 1
                
            Case "xls"
                GetFileType = 2
            Case "xlsm"
                GetFileType = 2
            Case "xlsx"
                GetFileType = 2
            Case "xlsb"
                GetFileType = 2
            Case "csv"
                GetFileType = 2
                
            Case "doc"
                GetFileType = 3
            Case "docx"
                GetFileType = 3
            Case "rtf"
                GetFileType = 3
                
            Case "txt"
                GetFileType = 4
                
            Case "ppt"
                GetFileType = 5
            Case "pptx"
                GetFileType = 5
                
            Case "jpg"
                GetFileType = 6
            Case "jpeg"
                GetFileType = 6
            Case "tif"
                GetFileType = 6
            Case "png"
                GetFileType = 6
                
            Case "msg"
                GetFileType = 8
                
            Case "zip"
                GetFileType = 10
            Case "rar"
                GetFileType = 10
                
            Case "ods"
                GetFileType = 20
            Case "odt"
                GetFileType = 21
            Case "sxw"
                GetFileType = 22
                
            ' unknown type is 7 !!!!!
            Case Else
                If Not FileExists(sfn) Then
                    GetFileType = 7
                Else
                    Dim f As File: Set f = FSO.GetFile(sfn)
                    If f.Size = 0 Then
                        GetFileType = 9
                    Else
                        GetFileType = 7
                    End If
                End If
        End Select
    End If
End Function


'*************************************
' load text file
'*************************************

Function LoadFileToCollection(sfn As String) As Collection
    Debug.Assert FileExists(sfn)
    
    Set LoadFileToCollection = New Collection
    Open sfn For Input As #1
    Do Until EOF(1)
        Dim sFileline As String: Line Input #1, sFileline
        LoadFileToCollection.Add sFileline
    Loop
    Close #1
End Function


Function LoadFileToArray(sfn As String) As Variant()
    Debug.Assert FileExists(sfn)
    Dim c As Collection: Set c = LoadFileToCollection(sfn)
    If c.Count = 0 Then c.Add ""
    LoadFileToArray = AutoOrientVector(CollectionToArray(c))
End Function


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



'>> Forms

Function FirstDropDown() As DropDown
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        If TypeOf shp.OLEFormat.Object Is DropDown Then
            Set FirstDropDown = shp.OLEFormat.Object
            Exit Function
        End If
    Next
End Function

Function NamedDropDown(Optional sDropDownName As String) As DropDown
    On Error GoTo done
    If sDropDownName = "" Then
        Set NamedDropDown = FirstDropDown
    Else
        Dim shp As Shape: Set shp = ActiveSheet.Shapes(sDropDownName)
        Set NamedDropDown = shp.OLEFormat.Object
    End If
done:
End Function



'>> Finance

Function CAGR(vLast As Variant, vFirst As Variant, Optional nYears As Integer) As Double
    Dim dLast As Double: dLast = AsDouble(vLast)
    Dim dFirst As Double: dFirst = AsDouble(vFirst)
    Dim State As Integer
    If nYears = 0 Then
        If (TypeOf vLast Is Range) And (TypeOf vFirst Is Range) Then
            nYears = vLast.Column - vFirst.Column
        End If
    End If
    
    If nYears = 0 Then
        CAGR = 0
    Else
        CAGR = (dLast / dFirst) ^ (1 / nYears) - 1
    End If
End Function


Function EquivIRR(dStart As Double, dEnd As Double, dYears As Double) As Double
    EquivIRR = (dEnd / dStart) ^ (1 / dYears) - 1
End Function



'>> Errors

'*************************************
' errors
'*************************************

Function IfError(v, vOnError) As Variant
    On Error GoTo err
    If IsError(v) Then IfError = vOnError Else IfError = v
    Exit Function
err:
    IfError = vOnError
End Function

Function ErrNA() As Variant
    ErrNA = CVErr(xlErrNA)
End Function

Function ErrValue() As Variant
    ErrValue = CVErr(xlErrValue)
End Function

Function VBError()
    'SoliError = "Error # " & Str(err.Number) & " was generated by " & err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & err.Description
    VBError = err.Description
    If err.Description <> "" Then VBError = "#" & VBError
End Function

Sub ShowVBError()
    If err.Number <> 0 Then
        MsgBox VBError(), , "Error", err.HelpFile, err.HelpContext
    End If
End Sub

Public Sub RaiseVBError(sMsg As String)
    err.Raise 99, , sMsg
End Sub

Public Sub Assert(b As Boolean, sMsg As String)
    If Not b Then RaiseVBError sMsg
End Sub


Sub MsgError(sError As String, Optional iErrorCode As Integer)
    MsgBox sError, vbCritical + vbOKOnly
    err.Raise iErrorCode
End Sub



'>> Helpers

' ((NNBQPDT)) comments in Helpers -> Comments
 
' change Excel undo history up to 100
'    https://support.microsoft.com/en-us/kb/211922

' refactor: mix helpers into Util(2)

'*************************************
' fiddling with Excel caption
'*************************************

Sub TestExcelCaption()
    TagExcelCaption "awert"
    MsgBox GetTagFromExcelCaption()
End Sub

Function IsValidExcelCaptionTag(sTag As String) As Boolean
    IsValidExcelCaptionTag = RegMatch(sTag, "^[A-Za-z0-9 ,.:-]+$")
End Function

Sub ClearExcelCaption()
    ActiveWindow.Caption = ActiveWorkbook.Name
End Sub

Sub TagExcelCaption(sTag As String)
    If IsValidExcelCaptionTag(sTag) Then
        ActiveWindow.Caption = FS("# (""#"")", ActiveWorkbook.Name, sTag)
    End If
End Sub

Function GetTagFromExcelCaption() As String
    Application.Volatile
    GetTagFromExcelCaption = RegMatch(ActiveWindow.Caption, """([A-Za-z0-9 ,.:-]+)""", 1)
End Function


'*************************************
' stuff
'*************************************

Function ref(ParamArray a())
    Dim s As String
    Dim iParam As Integer
    For iParam = LBound(a) To UBound(a)
        If iParam > LBound(a) Then s = s & ", "
        s = s & a(iParam)
    Next
    ref = s
End Function


Function LooksEmpty(r As Range) As Boolean
    LooksEmpty = IsEmpty(r) Or r.Value2 = ""
End Function


Sub DoReplace(sWhat As String, sReplacement As String)
    Cells.Replace what:=sWhat, replacement:=sReplacement, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub


Sub SafeSelect(r As Range)
    SetSilentApplicationState
    On Error GoTo done
    Dim rActive As Worksheet: Set rActive = ActiveSheet
    r.Worksheet.Activate
    r.Select
    rActive.Activate
done:
    RevertApplicationState
End Sub


Sub SafePasteSpecial(rTarget As Range, Optional pt As XlPasteType, Optional op As XlPasteSpecialOperation, Optional bSkipBlanks As Boolean, Optional bTranspose As Boolean)
    ' https://groups.google.com/forum/#!topic/microsoft.public.excel.worksheet.functions/RqjeigvN8oI
    Dim old: old = Application.ScreenUpdating
    On Error GoTo done
    If pt = 0 Then pt = xlPasteAll
    If op = 0 Then op = xlPasteSpecialOperationNone
    rTarget.PasteSpecial Paste:=pt, Operation:=op, SkipBlanks:=bSkipBlanks, Transpose:=bTranspose
done:
    Application.ScreenUpdating = old
End Sub


'*************************************
' AHK
'*************************************

Sub AutoFitRowHeight()
    ' ^!h
    ActiveCell.EntireRow.AutoFit
End Sub


Sub PutStrandInCell()
    ActiveCell.Value2 = RandomString(4)
End Sub


'*************************************
' Excel dependency matrix magic
'*************************************

Sub Touch(r As Range)
    ' scheint ganz einfach zu funktioniert
    'Dim vOld: vOld = r.Value
    'r.Value = vOld
    If r.HasFormula Then
        r.formula = r.formula
    Else
        r.value = r.value
    End If
End Sub


Function Link(ParamArray asdf()) As Boolean
    Link = True
    On Error Resume Next
    Link = Range("UseFunctions")
End Function


Function Link2(rTouchpoint As Range, func) As Variant
    Link2 = func
End Function

Function Link3(func, ParamArray rTouchpoints()) As Variant
    Link3 = func
End Function


Function LinkFiles(ParamArray asdf()) As Integer
    LinkFiles = UBound(asdf) - LBound(asdf) + 1
End Function


'*************************************
' cell flashing
'*************************************

Sub FlashCell(r As Range, iSleep As Integer)
    Dim clInterior As Long: clInterior = r.Interior.Color
    Dim clFont As Long: clFont = r.Font.Color
    r.Interior.Color = RGB(255, 0, 0)
    r.Font.Color = RGB(255, 255, 255)
    Sleep iSleep
    r.Interior.Color = clInterior
    r.Font.Color = clFont
End Sub


Sub ChangeDefaultFont()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ixStyle As Integer
    For ixStyle = 1 To wb.Styles.Count
        Dim f As Font: Set f = wb.Styles(ixStyle).Font
        If f.Name = "Calibri" Then f.Name = "Arial"
        If f.Size = 11 Then f.Size = 8
    Next ixStyle
End Sub


'*************************************
' scrubbing
'*************************************

Sub DoRemoveEmptyAndZero(r As Range, bRemoveErrors As Boolean)
    SetSilentApplicationState
    On Error GoTo done
    
    Dim rCell As Range
    For Each rCell In r
        If IsError(rCell.Value2) Then
            If bRemoveErrors Then rCell.ClearContents
        Else
            If Trim(rCell.Value2) = "" Or rCell.Value2 = 0 Then
                rCell.ClearContents
            End If
        End If
    Next
    
done:
    RevertApplicationState
End Sub


Sub RemoveEmptyAndZero()
    DoRemoveEmptyAndZero Selection, True
End Sub



'>> JumpStation

'*************************************
' init
'*************************************

Function InitJumpStation() As UserFormSelector
    Set InitJumpStation = UserFormSelector.CreateWithoutOptions(m_frmSelector)
    
    ' Initialize jump targets
    Set m_jumpTargets = New Dictionary
    Set m_sheetSpecificTargets = New Dictionary
End Function


'*************************************
' helpers
'*************************************

Public Sub EnsureSheetIsVisible(ws As Worksheet)
    ' Check if the sheet is hidden or very hidden and make it visible
    If ws.Visible = xlSheetVeryHidden Then
        ws.Visible = xlSheetVisible
    ElseIf ws.Visible = xlSheetHidden Then
        ws.Visible = xlSheetVisible
    End If
End Sub


'*************************************
' jump stack
'*************************************

Public Sub SaveSheetToJumpStack(wsSheetToSave As Worksheet)
    If wsSheetToSave Is Nothing Then Set wsSheetToSave = ActiveSheet
    If m_cBackSheets Is Nothing Then Set m_cBackSheets = New Collection
    m_cBackSheets.Add wsSheetToSave
End Sub

Public Sub JumpBack()
    If m_cBackSheets Is Nothing Then GoTo no_sheet
    If m_cBackSheets.Count = 0 Then GoTo no_sheet
    Dim wsBack As Worksheet: Set wsBack = m_cBackSheets(m_cBackSheets.Count)
    EnsureSheetIsVisible wsBack
    wsBack.Activate
    m_cBackSheets.Remove m_cBackSheets.Count
    Exit Sub
no_sheet:
    MsgBox "no sheet to jump back to", vbExclamation
End Sub


'*************************************
' jump targets
'*************************************

Public Sub AddJumpTarget(requiredSheet As String, ByVal Text As String, ByVal subName As String)
    If requiredSheet = "" Then
        m_jumpTargets.Add Text, subName
    Else
        m_sheetSpecificTargets.Add Text, requiredSheet & "|" & subName
    End If
End Sub


Public Sub AddEmptyLine()
    Dim sEmptyLine As String
    Do
        sEmptyLine = sEmptyLine & " "
        If Not m_jumpTargets.Exists(sEmptyLine) Then Exit Do
    Loop Until False
    AddJumpTarget "", sEmptyLine, ""
End Sub


Sub AddJumpTargetsKeys()
    ' call this after you have added all jump targets
    ' Add visible options to form
    Dim key As Variant
    
    ' Add global targets
    For Each key In m_jumpTargets.Keys
        m_frmSelector.AddOption AsString(key)
    Next key
    
    ' Add sheet-specific targets
    For Each key In m_sheetSpecificTargets.Keys
        Dim parts As Variant
        parts = Split(m_sheetSpecificTargets(key), "|")
        If LCase(ActiveSheet.Name) = LCase(parts(0)) Then
            m_frmSelector.AddOption AsString(key)
        End If
    Next key
End Sub

Private Function locExecuteTarget(v)
    If SheetExists(v) Then
        SaveSheetToJumpStack ActiveSheet
        Dim ws As Worksheet: Set ws = Worksheets(v)
        EnsureSheetIsVisible ws
        ws.Activate
    Else
        Application.Run v
    End If
End Function

Public Function ExecuteSelectedTarget()
    ' Execute selected target
    If m_jumpTargets.Exists(m_frmSelector.Selection) Then
        locExecuteTarget m_jumpTargets(m_frmSelector.Selection)
    ElseIf m_sheetSpecificTargets.Exists(m_frmSelector.Selection) Then
        Dim selectedParts As Variant
        selectedParts = Split(m_sheetSpecificTargets(m_frmSelector.Selection), "|")
        locExecuteTarget selectedParts(1)
    Else
        MsgBox "?"
    End If
End Function


'*************************************
' jump to linked address
'*************************************

Public Sub JumpToAddressOnClipboard()
    Dim sLinkedAddress As String: sLinkedAddress = Trim(GetFromClipboard())
    JumpToLinkedAddress sLinkedAddress
End Sub


Public Sub JumpToAddressInComment()
    Dim sComment As String: sComment = GetComment(Selection.Cells(1))
    Dim cCommentLines As Collection: Set cCommentLines = StringToCollection(sComment, Chr(13) & Chr(10))
    If cCommentLines.Count = 0 Then GoTo no_link
    Dim sLinkedAddress As String: sLinkedAddress = Trim(cCommentLines(1))
    If sLinkedAddress = "" Then GoTo no_link
    JumpToLinkedAddress sLinkedAddress
    Exit Sub
no_link:
    MsgBox "no LINK in first line of comment", vbExclamation
End Sub


Public Sub JumpToLinkedAddress(sLinkedAddress As String)
    Dim wsBack As Worksheet: Set wsBack = ActiveSheet
    
    Dim sAddress As String: sAddress = Trim(sLinkedAddress)
    
    ' ((YTNNBMD))
    sAddress = RegMatch(sAddress, "^([Ll][Ii][Nn][Kk]:?)? *=? *(.+)", 2)
    Dim sSheet As String: sSheet = RegMatch(sAddress, "^'(\[.+?\])?([^']+?)'!(.+$)", 2)
    If Not SheetExists(sSheet) Then GoTo err
    Dim sLocalAddress As String: sLocalAddress = RegMatch(sAddress, "^'(\[.+?\])?([^']+?)'!(.+$)", 3)
    
    On Error GoTo err
    Dim wsJumpTarget As Worksheet: Set wsJumpTarget = Worksheets(sSheet)
    EnsureSheetIsVisible wsJumpTarget
    wsJumpTarget.Activate
    Range(sLocalAddress).Select
    Application.GoTo Reference:=Range(sLocalAddress), Scroll:=True
    On Error GoTo 0
    
    SaveSheetToJumpStack wsBack
    GoTo done
err:
    MsgBox "no address on clipboard", vbExclamation
    wsBack.Activate
done:
End Sub



'>> Macros

Function TryRunMacro(sWorkbookName As String, sSub As String) As Boolean
    TryRunMacro = False
    
    ' Passed workbook wouldn't have any subs to call
    If EndsWith(sWorkbookName, ".xlsx") Then Exit Function

    ' We can only call subs in macro sheets
    If Not Contains(sWorkbookName, ".") Then sWorkbookName = sWorkbookName & ".xlsm"
    
    Dim bInWorkbook As Boolean: bInWorkbook = ActiveWorkbook.Name = sWorkbookName
    If bInWorkbook Then
        On Error GoTo done
        Application.Run sWorkbookName & "!" & sSub
        
        ' no error
        TryRunMacro = True
    Else
        ' do nothing
    End If
done:
End Function


Sub CtrlD()
    DoAllDependents
End Sub

Sub CtrlE()
    JumpToNextSheetInCircle
End Sub

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

Sub CtrlM()
    GotoNextPart
End Sub

Sub CtrlN()
    GotoPrevPart
End Sub

Sub CtrlP()
    'If Not TryRunMacro("Vertragsdatenbank.xlsm", "OpenInvoicePDFForActiveCell") Then
        DoAllPrecedents
    'End If
End Sub

Sub CtrlQ()
    ToggleActiveSheetInCircle
End Sub

Sub CtrlR()
Attribute CtrlR.VB_ProcData.VB_Invoke_Func = "r\n14"
    RenameSheet
End Sub

Sub CtrlW()
    JumpToPrevSheetInCircle
End Sub



'>> Names

' ((FFZPMWT)) remove DefineNamedRangeOld
' ((DANDPHZ)) remove FormatDottedBox

Sub UnhideAllNames()
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        n.Visible = True
    Next
End Sub


'*************************************
' newer stuff (Purobike)
'*************************************

Function SafeGetName(cell As Range) As Name
    Dim n As Name
    Dim rng As Range
    
    For Each n In ThisWorkbook.Names
        On Error Resume Next
        Set rng = n.RefersToRange
        On Error GoTo 0
        
        If Not rng Is Nothing Then
            If rng.Address(External:=True) = cell.Address(External:=True) Then
                Set SafeGetName = n
                Exit Function
            End If
        End If
        Set rng = Nothing
    Next n
    
    Set SafeGetName = Nothing
End Function


Function ParseExcelAddress(sAddress) As Variant
    Dim regEx As Object, matches As Object, match As Variant
    Dim result(1 To 3) As String
    
    ' Initialize regex
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .PATTERN = "(?:'?\[(.*?)\])?(?:'?(.*?)'?\!)?(.+)"
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
    End With

    ' Test the regex against the sAddress
    Set matches = regEx.Execute(sAddress)
    
    ' If matches are found, assign to result array
    If matches.Count > 0 Then
        Set match = matches(0)
        result(1) = match.SubMatches(0) ' Workbook
        result(2) = match.SubMatches(1) ' Sheet
        result(3) = match.SubMatches(2) ' Local sAddress
    End If
    
    ' Return result
    ParseExcelAddress = result
End Function


Function IsRangeName(nm As Name) As Boolean
    On Error Resume Next
    Dim rng As Range
    Set rng = nm.RefersToRange
    IsRangeName = (err.Number = 0)
    err.Clear
End Function

Function GetRangeName(rng As Range) As String
    Dim nm As Name

    ' Workbook-level names
    For Each nm In rng.Parent.Parent.Names
        If IsRangeName(nm) Then
            If nm.RefersToRange.Address = rng.Address _
               And nm.RefersToRange.Parent Is rng.Parent Then
                GetRangeName = nm.Name
                Exit Function
            End If
        End If
    Next nm

    ' Worksheet-level names
    For Each nm In rng.Parent.Names
        If IsRangeName(nm) Then
            If nm.RefersToRange.Address = rng.Address Then
                GetRangeName = nm.Name
                Exit Function
            End If
        End If
    Next nm
End Function


'*************************************
' marking
'*************************************

Sub MarkSelectionAsNamedRange()
    MarkAsNamedRange Selection
End Sub


Sub MarkAsNamedRange(rNamedRange As Range)
    
    Dim vEdge
    For Each vEdge In CreateCollection(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With rNamedRange.Borders(vEdge)
            .LineStyle = xlDash
            .Color = RGB(230, 184, 183)
            '.ThemeColor = 6
            '.TintAndShade = 0.599963377788629
            .Weight = xlMedium
        End With
    Next
    
    Dim vBorder
    For Each vBorder In CreateCollection(xlDiagonalDown, xlDiagonalUp, xlInsideVertical, xlInsideHorizontal)
        rNamedRange.Borders(vBorder).LineStyle = xlNone
    Next
End Sub


' old style, deprecated
' ((DANDPHZ)) remove FormatDottedBox
Public Sub FormatDottedBox()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub


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


Function TypedRange(ws As Worksheet, ByVal sType, ByVal sItem, sName) As Range
    sType = RegMatch(sType, "^[0-9]+ (.+)$", 1)
    sItem = RegMatch(sItem, "^([A-Za-z0-9]+) ?", 1)
    Dim sRangeName: sRangeName = FS("#_#_#", sType, sItem, sName)
    Set TypedRange = ws.Range(sRangeName)
End Function


'*************************************
' define
'*************************************

' 15.04.2016 somehow we don't need the check if name exists anymore ...
Function DefineNamedRange(sNamedRange As String, rNamedRange As Range, Optional bLocal As Boolean, Optional bDoNotMark As Boolean) As Range
    If bLocal Then
        Dim ws As Worksheet: Set ws = rNamedRange.Worksheet
        ws.Names.Add sNamedRange, rNamedRange
    Else
        ActiveWorkbook.Names.Add sNamedRange, rNamedRange
    End If

    If Not bDoNotMark Then MarkAsNamedRange rNamedRange
    
    Set DefineNamedRange = rNamedRange
End Function



'>> Params

'*************************************
' combining parameterized snippets
'*************************************

Function BulletList(ParamArray pas())
    Dim vParams(): vParams = ParamArrayDelegated(pas)
    Dim c As New Collection
    Dim ix
    For ix = LBound(vParams) To UBound(vParams)
        If TypeOf vParams(ix) Is Range Then
            Dim rv As Range: Set rv = vParams(ix)
            Dim r As Range
            For Each r In rv
                c.Add "* " & r.Value2
            Next
        Else
            Dim s As String: s = vParams(ix)
            c.Add "* " & s
        End If
    Next
    BulletList = CollectionToString(c, Chr(13) & Chr(10))
End Function


'*************************************
' formatting
'*************************************

Function SlideParams(ParamArray pas())
    Dim v: v = ParamArrayDelegated(pas)
    SlideParams = ConcatenateWithSep(v, " | ")
End Function


'*************************************
' formatting
'*************************************

' ****** old

Function FormatSentence(sSentence As String, v1 As Range, Optional v2 As Range, Optional v3 As Range, Optional v4 As Range, Optional v5 As Range, Optional v6 As Range) As String
    Const PARAMCOUNT = 6
    
    Dim ixParam As Integer
    Dim ix As Integer
    Do
        ix = InStr(1, sSentence, "#")
        Dim bVariableFound As Boolean: bVariableFound = ix <> 0
        If bVariableFound Then
        
            Dim sReplacement As String
            
            ixParam = ixParam + 1
            If ixParam > PARAMCOUNT Then
            
                FormatSentence = "# missing parameters"
                Exit Function
            
            End If
            
            Dim rParam As Range: Set rParam = GetParam(ixParam, v1, v2, v3, v4, v5, v6)
            ' http://stackoverflow.com/questions/6932901/how-do-i-retrieve-an-excel-cell-value-in-vba-as-formatted-in-the-worksheet
            sReplacement = Trim(rParam.Text)
        
            sSentence = Left(sSentence, ix - 1) & sReplacement & Mid(sSentence, ix + 1)
        
        End If
    Loop While bVariableFound
    
    FormatSentence = sSentence
    
End Function


Function FormatSentenceVolatile(sSentence As String, v1 As Range, Optional v2 As Range, Optional v3 As Range, Optional v4 As Range, Optional v5 As Range, Optional v6 As Range) As String
    Application.Volatile
    FormatSentenceVolatile = FormatSentence(sSentence, v1, v2, v3, v4, v5, v6)
End Function


'Function FS(sSentence As String, v1 As Range, Optional v2 As Range, Optional v3 As Range, Optional v4 As Range, Optional v5 As Range, Optional v6 As Range) As String
'    FS = FormatSentence(sSentence, v1, v2, v3, v4, v5, v6)
'End Function


' ****** new

Function IntInc(ByRef ix As Variant, Optional start As Integer) As Integer
    If IsEmpty(ix) Then ix = start
    IntInc = ix
    ix = ix + 1
End Function


Function StringifyParams(arrParams() As Variant) As Variant
    ' how to call, w/ "tokens" a ParamArray
    '     Dim arrParams(): arrParams = tokens: arrParams = StringifyParams(arrParams)

    ' passed params come from a ParamArray which is 0-based
    Debug.Assert LBound(arrParams) = 0
    
    Dim ixParam
    Dim r As Range
    
    If ArrayCount(arrParams) = 1 Then
        If TypeOf arrParams(LBound(arrParams)) Is Range Then
            Dim rSingleParam As Range: Set rSingleParam = arrParams(LBound(arrParams))
            Dim nCells As Integer: nCells = rSingleParam.Cells.Count
            If nCells <> 1 Then
                Dim expandedParams(): ReDim expandedParams(0 To nCells - 1)
                
                For Each r In rSingleParam.Cells
                    Set expandedParams(IntInc(ixParam, 0)) = r
                Next
                
                arrParams = expandedParams
            End If
        End If
    End If
    
    For ixParam = LBound(arrParams) To UBound(arrParams)
        If TypeOf arrParams(ixParam) Is Range Then
            Set r = arrParams(ixParam)
            ' http://stackoverflow.com/questions/6932901/how-do-i-retrieve-an-excel-cell-value-in-vba-as-formatted-in-the-worksheet
            arrParams(ixParam) = r.Text
        Else
            Dim sParam As String: sParam = arrParams(ixParam)
            arrParams(ixParam) = sParam
        End If
    Next
    
    StringifyParams = arrParams
End Function


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


Function FormatSentence2(sSentence As String, ParamArray tokens() As Variant) As String
    Application.Volatile
    Dim arrParams(): arrParams = tokens: arrParams = StringifyParams(arrParams)
    FormatSentence2 = DoFormatSentence2(sSentence, arrParams)
End Function

Function FS(sSentence As String, ParamArray tokens() As Variant) As String
    Dim arrParams(): arrParams = tokens: arrParams = StringifyParams(arrParams)
    FS = DoFormatSentence2(sSentence, arrParams)
End Function


' ****** new

Public Function SubstituteParams(mask As String, ParamArray tokens()) As String
    ' http://stackoverflow.com/questions/17233701/is-there-an-equivalent-of-printf-or-string-format-in-excel
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = Replace$(mask, "{" & i + 1 & "}", tokens(i))
    Next
    SubstituteParams = mask
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


Function GetParam(ix As Integer, v1 As Range, Optional v2 As Range, Optional v3 As Range, Optional v4 As Range, Optional v5 As Range, Optional v6 As Range) As Range
    Select Case ix
        Case 1
            Set GetParam = v1
        Case 2
            Set GetParam = v2
        Case 3
            Set GetParam = v3
        Case 4
            Set GetParam = v4
        Case 5
            Set GetParam = v5
        Case 6
            Set GetParam = v6
    End Select
End Function



' >> PDF

Function GetWorksheetNameFromRef(rDirect As Range) As String
    ' =ExtractPdfNumber('2012 P&L Details'!$C$5)
    Const RE1 = "=-?ExtractPdfNumber\(([^)]+)\)"
    Dim sRef As String: sRef = Replace(RegMatch(rDirect.formula, RE1, 1), "'", "")
    If sRef = "" Then
        ' don't do anything
        GetWorksheetNameFromRef = ""
    Else
        Const RE2 = "([^!]+)!(.+)"
        Dim sDirectSheet  As String: sDirectSheet = RegMatch(sRef, RE2, 1)
        GetWorksheetNameFromRef = sDirectSheet
    End If

End Function

Function GetMatchFromRef(rDirect As Range) As String
    ' =ExtractPdfNumber('2012 P&L Details'!$C$5)
    Const RE1 = "=-?ExtractPdfNumber\(([^)]+)\)"
    Dim sRef As String: sRef = Replace(RegMatch(rDirect.formula, RE1, 1), "'", "")
    If sRef = "" Then
        ' don't do anything
        GetMatchFromRef = ""
    Else
        Const RE2 = "([^!]+)!(.+)"
        Dim sDirectSheet  As String: sDirectSheet = RegMatch(sRef, RE2, 1)
        Dim sDirectAddress As String: sDirectAddress = RegMatch(sRef, RE2, 2)
        Dim rDirectRef As Range: Set rDirectRef = Worksheets(sDirectSheet).Range(sDirectAddress)
        Dim rMatch As Range: Set rMatch = Worksheets(sDirectSheet).Columns("A:A")
        GetMatchFromRef = rMatch.Cells(rDirectRef.Row)
    End If
End Function


Sub MakeAsDouble()
    Dim r As Range
    For Each r In Selection.Cells
        If Left(r.formula, 1) = "=" Then
            r.formula = "=ConvertPdfNumber(" & Mid(r.formula, 2) & ")"
        End If
    Next
End Sub

Function ConvertPdfNumber(v) As Double
    ConvertPdfNumber = AsDouble(v)
End Function


Function ExtractPdfNumber(v, Optional ix As Integer) As Double
    If v = "-" Or v = "" Then
        ExtractPdfNumber = 0
    Else
        If ix = 0 Then ix = 1
        ExtractPdfNumber = RegMatch(v, "([0-9.,-]+)( +([0-9..-]+))?", ix)
    End If
End Function



' >> Pivot

' http://stackoverflow.com/questions/12861909/how-to-create-a-pivot-table-in-vba
' http://www.mrexcel.com/forum/excel-questions/711451-tabledestination-pivot-table-macro.html
' http://www.thespreadsheetguru.com/blog/2014/9/27/vba-guide-excel-pivot-tables
' http://analysistabs.com/excel-vba/pivot-tables-examples/


'*************************************
' helpers
'*************************************

Function GetPivotTableDimensionsRange(pt As PivotTable) As Range
    Dim rUpperLeft As Range: Set rUpperLeft = pt.TableRange1.Cells(1).Offset(2, 0)
    Dim rLowerRight As Range: Set rLowerRight = rUpperLeft.End(xlDown).Offset(-1, 1)
    Set GetPivotTableDimensionsRange = rUpperLeft.Worksheet.Range(rUpperLeft, rLowerRight)
End Function


Function GetPivotTableDataRange(pt As PivotTable, nPivotHeaderRows As Integer) As Range
    Dim wsPivot As Worksheet: Set wsPivot = pt.Parent
    With pt.TableRange1
        Dim rTopLeft As Range: Set rTopLeft = .Cells(nPivotHeaderRows + 1, 1)
        Dim rBottomRight As Range: Set rBottomRight = .Cells(.Rows.Count, .Columns.Count)
    End With
    Set GetPivotTableDataRange = wsPivot.Range(rTopLeft, rBottomRight)
End Function


'*************************************
' copy values from pivot to data sheet
'*************************************

' deprecated, now using CopyColumns()

'Sub CopyPivotValues(pt As PivotTable, nPivotHeaderRows As Integer, rTargetAnchor As Range, Optional nPivotCols As Integer)
'    Dim rTopLeft As Range: Set rTopLeft = pt.TableRange1.Cells(1).Offset(nPivotHeaderRows, 0)
'
'    Dim rBottomRight As Range
'    If nPivotCols = 0 Then
'        ' will copy the complete pivot range
'        Set rBottomRight = rTopLeft.End(xlDown).End(xlToRight)
'    Else
'        ' only copy passed number of columns
'        Set rBottomRight = rTopLeft.End(xlDown).Offset(0, nPivotCols - 1)
'    End If
'
'    Dim rPivotData As Range: Set rPivotData = rTopLeft.Worksheet.Range(rTopLeft, rBottomRight)
'
'    Dim wsTarget As Worksheet: Set wsTarget = rTargetAnchor.Worksheet
'
'    ' Pivot values will go into the left columns of the target table
'    ' target table will usually have more columns = calculated fields, based on some of the pivot data
'    Dim rTargetTable As Range: Set rTargetTable = wsTarget.Range(rTargetAnchor, rTargetAnchor.End(xlDown).End(xlToRight))
'
'    Dim nNewRows As Integer: nNewRows = rPivotData.Rows.Count - rTargetTable.Rows.Count
'
'    If nNewRows > 0 Then
'
'        ' new rows in pivot
'
'        ' copy the last row in the target table ...
'        Dim rLastTargetRow As Range: Set rLastTargetRow = rTargetTable.Rows(rTargetTable.Rows.Count)
'        rLastTargetRow.Copy
'
'        ' ... so that we have the same # of rows in the target as in the pivot
'        Dim rNewTargetRows As Range: Set rNewTargetRows = wsTarget.Range(rLastTargetRow.Cells(1, 1).Offset(1), rLastTargetRow.Offset(nNewRows)).EntireRow
'
'        ' somehow we can't paste directly to the row range, so iterate over the new target rows
'        Dim rRow As Range
'        For Each rRow In rNewTargetRows.Rows
'            rRow.PasteSpecial xlPasteAll
'        Next
'
'        Application.CutCopyMode = False
'    Else
'
'        If nNewRows < 0 Then
'            ' fewer rows in pivot than in target
'
'            ' NOTE: nNewRows is negative
'            Dim ixFirstTargetRowToClear As Integer: ixFirstTargetRowToClear = rTargetTable.Rows.Count + nNewRows + 1
'
'            Dim ixClear As Integer
'            For ixClear = ixFirstTargetRowToClear To rTargetTable.Rows.Count
'
'                ' intentionally only clear the contents, so that we can see that we have fewer table rows
'                rTargetTable.Rows(ixClear).ClearContents
'            Next
'        End If
'    End If
'
'    ' now exchange the existing data in the target w/ values from the pivot
'    rPivotData.Copy
'    wsTarget.Range("TargetAnchor").PasteSpecial xlPasteValues
'    Application.CutCopyMode = False
'End Sub


'*************************************
' PivotCache
'*************************************

Function CreatePivotCache(rSourceData As Range) As PivotCache
    Dim sSourceData As String: sSourceData = rSourceData.Address(ReferenceStyle:=xlR1C1)
    Set CreatePivotCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sSourceData, Version:=xlPivotTableVersion14)
End Function


Sub PivotCacheReport()
    ' http://ramblings.mcpher.com/Home/excelquirks/snippets/pivotcache
    Dim pc As PivotCache
    Dim s As String, sn As String
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    With ActiveWorkbook
        For Each pc In .PivotCaches
            Dim c As Collection
            Set c = New Collection
            
            's = "Pivotcache " & CStr(pc.index) & " uses " & CStr(pc.MemoryUsed) & " and has " _
            '        & CStr(pc.RecordCount) & " records"
            c.Add "Pivotcache " & CStr(pc.index) & " has " & CStr(pc.RecordCount) & " records"
            c.Add "The following pivot tables use it:"
            
            For Each ws In .Worksheets
                For Each pt In ws.PivotTables
                    If pt.CacheIndex = pc.index Then

                        c.Add "* '" & pt.Name & "' on " & ws.Name
                    End If
                Next pt
            Next ws
            
            MsgBox CollectionToString(c, Chr(13) & Chr(10))
        Next pc
        
        sn = Chr(10) & "Couldnt find the pivotcache for these pivot tables"
        s = ""
        For Each ws In .Worksheets
            For Each pt In ws.PivotTables
                If pt.CacheIndex < 1 Or pt.CacheIndex > .PivotCaches.Count Then
                    s = s & Chr(10) & ws.Name & ":" & Replace(pt.Name, "PivotTable", "PT")
                End If
            Next pt
        Next ws
        If (Len(s) > 0) Then
            MsgBox (sn & s)
        End If
    End With
End Sub


'*************************************
' PivotTable
'*************************************

Private Function GetRangeFromZS(sSourceData As String) As Range
    Dim a() As String: a = Split(sSourceData, "!")
    Dim sR1C1 As String: sR1C1 = a(0) & "!" & Replace(Replace(a(1), "Z", "R"), "S", "C")
    Dim sA1 As String: sA1 = Application.ConvertFormula(sR1C1, xlR1C1, xlA1)
    Set GetRangeFromZS = Range(sA1)
End Function


Function CreatePivot(sTableName As String, rSourceData As Range, rDestination As Range) As PivotTable
    Dim pc As PivotCache
    For Each pc In ActiveWorkbook.PivotCaches
        Dim rExistingSourceData As Range: Set rExistingSourceData = GetRangeFromZS(pc.SourceData)
        If rExistingSourceData.Address(External:=True) = rSourceData.Address(External:=True) Then
            Exit For
        End If
    Next
    
    If pc Is Nothing Then
        Set pc = CreatePivotCache(rSourceData)
    End If
    
    Set CreatePivot = pc.CreatePivotTable(TableDestination:=rDestination, TableName:=sTableName, DefaultVersion:=xlPivotTableVersion14)
End Function


'*************************************
' adding fields
'*************************************

Sub AddField(pt As PivotTable, sFieldName As String, orient As XlPivotFieldOrientation)
    With pt.PivotFields(sFieldName)
        .Orientation = orient
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
End Sub

Sub AddPageField(pt As PivotTable, sFieldName As String)
    AddField pt, sFieldName, xlPageField
End Sub

Sub AddRowField(pt As PivotTable, sFieldName As String)
    AddField pt, sFieldName, xlRowField
End Sub

Sub AddColumnField(pt As PivotTable, sFieldName As String)
    AddField pt, sFieldName, xlColumnField
End Sub

Function FunctionToString(func As Variant)
    Select Case func
        Case xlAverage
            FunctionToString = "Average"
        Case xlCountNums
            FunctionToString = "CountNums"
        Case xlMin
            FunctionToString = "Min"
        Case xlStDev
            FunctionToString = "StDev"
        Case xlSum
            FunctionToString = "Sum"
        Case xlVar
            FunctionToString = "Var"
        Case xlCount
            FunctionToString = "Count"
        Case xlMax
            FunctionToString = "Max"
        Case xlProduct
            FunctionToString = "Product"
        Case xlStDevP
            FunctionToString = "StDevP"
        Case xlVarP
            FunctionToString = "VarP"
        Case Else
            FunctionToString = "?"
    End Select
End Function


Sub AddDataField(pt As PivotTable, sFieldName As String, func As Variant)
    ' xlAverage, xlCountNums, xlMin, xlStDev, xlSum, xlVar, xlCount, xlMax, xlProduct, xlStDevP, or xlVarP.
    Dim sCaption As String: sCaption = FunctionToString(func) & " of " & sFieldName
    pt.AddDataField pt.PivotFields(sFieldName), sCaption, func
End Sub


'*************************************
' sorting
'*************************************

Sub AutosortField(pt As PivotTable, sFieldName As String, so As XlSortOrder)
    pt.PivotFields(sFieldName).AutoSort so, sFieldName
End Sub


'*************************************
' filters
'*************************************

Function OnOff(aVisible()) As Dictionary
    Set OnOff = New Dictionary
    Dim ix As Integer
    For ix = LBound(aVisible) To UBound(aVisible)
        Dim sAttribute As String: sAttribute = aVisible(ix)
        If Not OnOff.Exists(sAttribute) Then OnOff.Add sAttribute, sAttribute
    Next ix
End Function


Sub CreateFilter(pt As PivotTable, sFieldName, ParamArray aVisible())
    Dim v(): v = aVisible
    Dim oo As Dictionary: Set oo = OnOff(v)
    
    With pt.PivotFields(sFieldName)
        .CurrentPage = "(All)"
        .EnableMultiplePageItems = True
        
        Dim pi As PivotItem
        For Each pi In .PivotItems
            pi.Visible = oo.Exists(pi.Name)
        Next
    End With
End Sub


'*************************************
' adjusting pivot data source
'*************************************

Sub AdjustPivotSourceData(pt As PivotTable, rNewSourceData As Range)
    ' see also http://www.thespreadsheetguru.com/the-code-vault/2014/7/9/change-a-pivot-tables-data-source-range
    
    ' Make sure every column in data set has a heading and is not blank (error prevention)
    If WorksheetFunction.CountBlank(rNewSourceData.Rows(1)) > 0 Then
        MsgError "One of your data columns has a blank heading." & vbNewLine & "Please fix and re-run!"
    End If
    
    ' passed SourceData must be string in R1C1 notation
    Dim sNewSourceData As String
    sNewSourceData = rNewSourceData.Worksheet.Name & "!" & rNewSourceData.Address(ReferenceStyle:=xlR1C1)
    
    ' Change Pivot Table Data Source Range Address
    pt.ChangePivotCache _
        ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=sNewSourceData)
          
    'Ensure Pivot Table is Refreshed
    pt.RefreshTable
    
    'Complete Message
    Application.StatusBar = pt.Name & "'s data source range has been successfully updated!"
End Sub



' >> Pomodoro

Function CountPoms(r As Range)
    Dim arr() As String: arr = Split(r.Value2, Chr(10))
    
    Dim nPoms
    
    Dim ixLine
    For ixLine = LBound(arr) To UBound(arr)
        Dim sLine: sLine = RegMatch(arr(ixLine), "^ *\* +(.+)$", 1)
        Dim sLineText: sLineText = RegMatch(sLine, PAT_LineParts, 1)
        Dim sPoms: sPoms = RegMatch(sLine, PAT_LineParts, 2)
        Dim iPoms As Integer: iPoms = IIf(sPoms = "", 0, sPoms)
        nPoms = nPoms + iPoms
    Next
    
    CountPoms = nPoms
End Function



'>> Powerpoint

'*************************************
' called from AHK
'*************************************

Public Sub TestPowerpointAvailable()

    ' check if we actually have a Powerpoint to call
    ForceInitPowerpoint
    If Not SafeInitPowerpoint Then Exit Sub
    
    ' show a MsgBox from the Powerpoint macro
    ' NOTE: this does not put Powerpoint in the foreground
    ' IMPORTANT: the ...!... format is necessary for the first call, afterwards it works w/ prefixing w/ ...!
    ' (reasons unknown, w/o the ...!... I get a powerpoint runtime error: -2147188160 invalid request sub or function not defined)
    ' note that no problem occures if we open the PowerPoint VBA window once - - after even the first non-prefixed Run() works fine
    On Error GoTo err
    g_Powerpoint.Run g_Powerpoint.ActivePresentation.Name & "!" & "TestPowerpointAvailableInPowerpoint"
    Exit Sub
    
err:
    ' if there never was a PowerPoint object, then the SafeInitPowerpoint wouldn't work
    ' because it worked, it probably means that there was a PowerPoint object, but PowerPoint was closed
    ' in this case, the reference in this Excel VBA is still there, albeit not valid anymore
    MsgBox "cannot call TestPowerpointAvailableInPowerpoint (PowerPoint object disappeared?) - might try again"
    
    ' prepare try again
    On Error Resume Next
    ForceInitPowerpoint
End Sub


Public Sub CopyAllShapesToPowerpoint()
    If Not SaveCurrentPowerpointSlide Then Exit Sub
    
    SetSilentApplicationState
    
    Dim vb: vb = MsgBox("All sheets (no = current)?", vbYesNoCancel)
    If vb = vbCancel Then GoTo done
    
    If vb = vbNo Then
        vb = MsgBox("Retain width?", vbYesNoCancel)
        If vb = vbCancel Then GoTo done
        locCopyAllShapesToPowerpoint ActiveSheet, vb = vbYes
    Else
        Dim sheet As Worksheet
        For Each sheet In Worksheets
            locCopyAllShapesToPowerpoint sheet, RETAIN_WIDTH
        Next
    End If
    
done:
    RevertApplicationState
    RestoreCurrentPowerpointSlide
End Sub


Sub locCopyAllShapesToPowerpoint(ws As Worksheet, bRetainWidth As Boolean)
    CopyAllShapesOnSheetToPowerpoint ws, bRetainWidth
    ReplaceSlideTexts ws
End Sub


Sub ReplaceSlideTexts(ws As Worksheet)
    If Not SafeInitPowerpoint Then GoTo done
    'PasteEnhancedMetafileFromExcel = False
    On Error GoTo done
    
    Dim dictPatternsNotFound As New Dictionary
    
    Const PATTERN = "^(.+?)( - text)? \(ppt\)$"
    Dim bPowerpointSheet: bPowerpointSheet = RegMatch(ws.Name, PATTERN)
    If Not bPowerpointSheet Then Exit Sub
    
    Dim rSlideTexts As Range: Set rSlideTexts = SafeGetNamedRange(R_SLIDE_TEXTS, ws)
    If Not rSlideTexts Is Nothing Then
    
        Dim iUseCol, iPatternCol, iGroupCol, iReplaceCol
        If rSlideTexts.Columns.Count = 3 Then
            iUseCol = -1: iPatternCol = 1: iGroupCol = 2: iReplaceCol = 3
        Else
            iUseCol = 1: iPatternCol = 2: iGroupCol = 3: iReplaceCol = 4
        End If
    
        Dim sSlideName As String: sSlideName = RegMatch(ws.Name, PATTERN, 1)
        
        Dim sPattern As String
        Dim sReplace As String
        Dim iGroup As Integer
    
        ' exach row in SlideTexts named range has a pattern and a replace value (or text snippet)
        Dim rRow As Range
        For Each rRow In rSlideTexts.Rows
        
            sPattern = rRow.Cells(iPatternCol)
            Dim bUse: bUse = (sPattern <> "") And ((iUseCol < 1) Or rRow.Cells(iUseCol))
            If bUse Then
            
                iGroup = rRow.Cells(iGroupCol)
                sReplace = rRow.Cells(iReplaceCol)
                
                If Contains(sReplace, "|") Then
                
                    Dim c As Collection: Set c = StringToCollection(sReplace, "|")
                    Dim ixReplace As Integer
                    For ixReplace = 1 To c.Count
                        Dim sReplacePart As String: sReplacePart = Trim(c(ixReplace))
                        Dim sThisPattern As String: sThisPattern = sPattern
                        Handle sSlideName, sThisPattern, ixReplace, sReplacePart, dictPatternsNotFound
                    Next
                
                Else
                
                    Handle sSlideName, sPattern, iGroup, sReplace, dictPatternsNotFound
                End If
                
            End If
        Next
    End If
    
done:
End Sub


Sub Handle(sSlideName As String, sPattern As String, iGroup As Integer, sReplace As String, dictPatternsNotFound As Dictionary)
    ' do the replace
    Dim nReplacements As Integer
    Dim vResult: vResult = SafePowerpoint.Run("FindReplaceAllOnSlideFromExcel", sSlideName, sPattern, iGroup, sReplace, nReplacements)
    
    Select Case nReplacements
        Case 0
            ' found but no need to change anything
        
        Case -1
            ' no match found
            If Not dictPatternsNotFound.Exists(sPattern) Then
            
                ' all other pattern which are equal to this one won't match either, so only show the msgbox once
                MsgBox FS("cannot find '#' for group #", sPattern, iGroup), vbExclamation + vbOKCancel
                dictPatternsNotFound.Add sPattern, sPattern
            End If
        
        Case 1
            ' no need to show a MsgBox, already done in Powerpoint
        
        Case Else
            MsgBox FS("# replacements for '#'!", sPattern), vbExclamation + vbOKCancel
    End Select
End Sub


'*************************************
' Powerpoint moniker
'*************************************

Sub ForceInitPowerpoint()
    Set g_Powerpoint = Nothing
    SafeInitPowerpoint
End Sub


Function SafePowerpoint() As Object
    SafeInitPowerpoint
    Set SafePowerpoint = g_Powerpoint
End Function


Function SafeInitPowerpoint() As Boolean
    If g_Powerpoint Is Nothing Then
    
        On Error GoTo err
        Set g_Powerpoint = GetObject(, "Powerpoint.Application")
        If g_Powerpoint Is Nothing Then GoTo err
        
        ' from here on we have a valid pointer to the Powerpoint object
        
        ' I don't remember why I put that here :)
        If g_Powerpoint.Version >= 9 Then g_Powerpoint.Visible = msoTrue
        
        ' below was a test, to prevent runtime error on 1st call w/o having opened the VBA window first
        ' => see fix in TestPowerpointAvailable
        'Dim sld As Object: Set sld = g_Powerpoint.ActiveWindow.View.Slide

    End If
    
    ' success
    SafeInitPowerpoint = True
    Exit Function
    
err:
    MsgBox "cannot GetObject Powerpoint.Application (SafeInitPowerpoint)"
    
    ' IMPORTANT: callers must check result before using the g_Powerpoint variable
    ' see e.g. TestPowerpointAvailable() below
    SafeInitPowerpoint = False
End Function


'*************************************
' called ThisWorkbook
'*************************************

' adjust rows on a powerpoint text sheet

Private Function IsEmptyRow(rRow As Range) As Boolean
    Dim rCell As Range
    For Each rCell In rRow
        If Not IsEmpty(rCell.Value2) Then
            IsEmptyRow = False
            Exit Function
        End If
    Next
    IsEmptyRow = True
End Function


' called from "DieseArbeitsmappe"
Public Sub Workbook_SheetChange_PptSheet(ByVal sh As Object, ByVal Target As Range)
    If EndsWith(sh.Name, " (ppt)") Then
        If NameExists("Text1") Then
            Dim r As Range: Set r = Range("Text1")
            Dim rIntersect As Range: Set rIntersect = Intersect(Target, r)
            If Not rIntersect Is Nothing Then
            
                Application.ScreenUpdating = False
                Application.Calculation = xlCalculationManual
                
                Dim rEntireRow As Range: Set rEntireRow = Target.EntireRow
                
                If IsEmptyRow(rIntersect) Then
                    rEntireRow.RowHeight = 3
                Else
                    rEntireRow.AutoFit
                    'rEntireRow.RowHeight = rEntireRow.RowHeight + 4
                End If
                
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
            
            End If
        End If
    End If
End Sub


'*************************************
' Powerpoint: save/restore slide
'*************************************

Private Function SaveCurrentPowerpointSlide() As Boolean
    SaveCurrentPowerpointSlide = SafeInitPowerpoint
    If SaveCurrentPowerpointSlide Then g_Powerpoint.Run "SaveCurrentSlide"
End Function


Private Function RestoreCurrentPowerpointSlide() As Boolean
    RestoreCurrentPowerpointSlide = SafeInitPowerpoint
    If RestoreCurrentPowerpointSlide Then g_Powerpoint.Run "RestoreCurrentSlide"
End Function


'*************************************
' copying shapes (text, chart)
'*************************************

Public Sub CopyNamedRangesOnSheetToPowerpoint()
    Const PATTERN = "^(.+?)( - text)? \(ppt\)$"
    Const TRAILING_PPT = " (ppt)"

    If Not SaveCurrentPowerpointSlide Then Exit Sub
    
    SetSilentApplicationState
    
    Dim sheet As Worksheet: Set sheet = ActiveSheet
    Dim bPowerpointSheet: bPowerpointSheet = RegMatch(sheet.Name, PATTERN)
        
    If Not bPowerpointSheet Then
    
        MsgBox sheet.Name & " is not a Powerpoint sheet"
        
    Else
        Dim sSlideName As String: sSlideName = RegMatch(sheet.Name, PATTERN, 1)
        
        Dim vb: vb = MsgBox("Retain width?", vbYesNoCancel)
        If vb = vbCancel Then GoTo done
        Dim bRetainWidth As Boolean: bRetainWidth = vb = vbYes
        
        Dim bSuccess As Boolean: bSuccess = CopySheetLocalNamedRangesToPowerpoint(sheet, sSlideName, bRetainWidth)
        If Not bSuccess Then
        
            MsgBox sheet.Name & "/" & sSlideName & ": problem to copy named ranges"
            
        End If
    End If
    
done:
    RevertApplicationState
    RestoreCurrentPowerpointSlide
End Sub


Sub CopyAllShapesOnSheetToPowerpoint(sheet As Worksheet, bRetainWidth As Boolean)
    ' Const TRAILING_PPT = " (ppt)"
    ' Dim bPowerpointSheet: bPowerpointSheet = EndsWith(sheet.Name, TRAILING_PPT)
    
    Const PATTERN = "^(.+?)( - text)? \(ppt\)$"
    Dim bPowerpointSheet: bPowerpointSheet = RegMatch(sheet.Name, PATTERN)
    
    If bPowerpointSheet Then
    
        'Dim sSlideName As String: sSlideName = Left(sheet.Name, Len(sheet.Name) - Len(TRAILING_PPT))
        Dim sSlideName As String: sSlideName = RegMatch(sheet.Name, PATTERN, 1)
        
        Dim bContinue
        
        ' for bulk upload to Powerpoint, always retain the width of the target shapes
        ' otherwise we would probably wreck the whole presentation :)
        
        bContinue = CopySheetLocalNamedRangesToPowerpoint(sheet, sSlideName, bRetainWidth)
        If Not bContinue Then GoTo done
        
        bContinue = CopyChartsOnSheetToPowerpoint(sheet, sSlideName, bRetainWidth)
        If Not bContinue Then GoTo done
        
    End If
    
done:
End Sub


Private Sub ParseNamedRangeName(n As Name, ByRef sSheet, ByRef sName As String, ByRef sAddress As String)
    Const PATTERN = "^=?'([^']+)'!(.+)$"
    sSheet = RegMatch(n.Name, PATTERN, 1)
    sName = RegMatch(n.Name, PATTERN, 2)
    sAddress = RegMatch(n.value, PATTERN, 2)
End Sub


Function CopySheetLocalNamedRangesToPowerpoint(sheet As Worksheet, sSlideName As String, bRetainWidth As Boolean) As Boolean
    CopySheetLocalNamedRangesToPowerpoint = sheet.Names.Count = 0
    If CopySheetLocalNamedRangesToPowerpoint Then Exit Function
    
    Dim nameOnSheet As Name
    For Each nameOnSheet In sheet.Names
        
        Dim sSheet As String
        Dim sName As String
        Dim sAddress As String
        ParseNamedRangeName nameOnSheet, sSheet, sName, sAddress
        
        'If sName = "Print_Area" Then
        If Not StartsWith(sName, "Table") Then
        
            ' never copy Druckbereich
            
        Else
        
            Dim sShapeName As String: sShapeName = sName
            
            On Error Resume Next
            Dim rSource As Range: Set rSource = sheet.Range(nameOnSheet.Name)
            On Error GoTo 0
            
            Debug.Assert Not rSource Is Nothing
                                    
            Dim ixRow As Integer
            If True Then
                ixRow = rSource.Rows.Count
            Else
                For ixRow = rSource.Rows.Count To 1 Step -1
                    If IsEmpty(rSource.Cells(ixRow, 1)) Then
                        ' still empty - skip
                    Else
                        ' encountered the last non-empty line
                        Exit For
                    End If
                Next
            End If
            
            With rSource
                Dim rCopy As Range: Set rCopy = .Worksheet.Range(.Cells(1, 1), .Cells(ixRow, .Columns.Count))
            End With
            
            'rCopy.Select
            
            'Dim bOk: bOk = CopyRangeToPowerpoint(rCopy, sSlideName, sShapeName)
            Dim bOk: bOk = PasteEnhancedMetafileFromExcel(rCopy, sSlideName, sShapeName, bRetainWidth)
            If Not bOk Then
                Dim yesno: yesno = MsgBox("cannot copy shape '" & sShapeName & "' on slide '" & sSlideName & "' - abort?", vbYesNo)
                If yesno = vbYes Then
                
                    ' maybe Powerpoint is not open, or not the correct Powerpoint
                    ' going through all the (ppt) slides doesn't make sense then, because all would result in error
                    Exit Function
                End If
            End If
        End If
    Next
    
    CopySheetLocalNamedRangesToPowerpoint = True
End Function


Function CopyChartsOnSheetToPowerpoint(sheet As Worksheet, sSlideName As String, bRetainWidth As Boolean) As Boolean
    CopyChartsOnSheetToPowerpoint = False
    
    Dim shp As Shape
    For Each shp In sheet.Shapes
        
        ' https://msdn.microsoft.com/en-us/library/aa432678(v=office.12).aspx
        Dim bShapeIsChart As Boolean: bShapeIsChart = shp.Type = msoChart
        If bShapeIsChart Then
            
            Dim sChartName As String: sChartName = shp.Name
            If Not StartsWith(sChartName, "Chart") Then
            
                ' do not try to copy just any chart
            
            Else
                'Dim bOk: bOk = CopyChartToPowerpoint(shp, sSlideName, sChartName, bRetainWidth)
                Dim bOk: bOk = PasteEnhancedMetafileFromExcel(shp, sSlideName, sChartName, bRetainWidth)
                If Not bOk Then
                
                    Dim yesno: yesno = MsgBox("cannot copy chart '" & sChartName & "' on slide '" & sSlideName & "' - abort?", vbYesNo)
                    If yesno = vbYes Then
                    
                        ' see above
                        Exit Function
                    End If
                
                End If
            End If
        End If
    Next
    CopyChartsOnSheetToPowerpoint = True
End Function


Sub TestCopyPicture()
    DoEvents
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
End Sub


Function PasteEnhancedMetafileFromExcel(rToCopy As Variant, sSlideName As String, sShapeName As String, bRetainWidth As Boolean) As Boolean
    If Not SafeInitPowerpoint Then GoTo done
    PasteEnhancedMetafileFromExcel = False
    On Error GoTo done
    
    If Contains(sShapeName, "IGNORE") Then
        PasteEnhancedMetafileFromExcel = True
    Else
        ' 02.01.2018
        ' used to be .Copy, replaced w/ CopyPicture (inspired by Rödl, from act3, see CopyPrintedPicture below)
        ' .CopyPicture seems to have a better kerning for Arial (so maybe we don't need Segoe UI anymore?)
        
        ' 05.02.25 this was what we had
        'rToCopy.Copy
        
        ' 05.02.25
        ' we cannot use xlPrinter, must be xlScreen
        ' note that we need a bitmap, because otherwise we have kerning issues in Powerpoint
        ' IMPORTANT: Wir müssen ScreenUpdating ausschalten, weil wir sonst beim CopyPicture nur eine weiße Fläche auf dem Clipboard haben
        
        Dim old: old = Application.ScreenUpdating
        Application.ScreenUpdating = True
        rToCopy.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        Application.ScreenUpdating = old
        
        'Application.CutCopyMode = False
        'rToCopy.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        'rToCopy.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        
        PasteEnhancedMetafileFromExcel = g_Powerpoint.Run("PasteEnhancedMetafileFromExcel", sSlideName, sShapeName, bRetainWidth)
    End If
done:
    Application.CutCopyMode = False
End Function


Sub CopyPrintedPicture()
    ' this is the sub inspired by Rödl, see PasteEnhancedMetafileFromExcel()
    Dim r As Range: Set r = Selection
    If NameExists("Druckbereich") Then
        If MsgBox("Use Druckbereich?", vbYesNoCancel) = vbYes Then Set r = Range("Druckbereich")
    End If
    r.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
End Sub


'*************************************
' F... formatting
'*************************************

' see also ((DFDCHYT)), Powerpoint

Function ReplaceSpace(s As String)
    ReplaceSpace = Replace(s, " ", ChrW(&H200A) & ChrW(&H200A) & ChrW(&H200A))
End Function

Function ReplaceMinus(s As String)
    ReplaceMinus = Replace(s, "-", ChrW(&H2212))
End Function

Function FNum(d As Double) As String
    FNum = ReplaceMinus(Format(d, "#,##0"))
End Function

Function FKm(d As Double, Optional nDigits As Integer = 0) As String
    ' &H2009: "kleiner Zwischenraum"
    If nDigits >= 1 Then
        FKm = FThousand(d, nDigits) & ChrW(&H2009) & "km"
    Else
        FKm = FThousand(d, nDigits) & ChrW(&H2009) & "km"
    End If
End Function

Function FPct(d As Double, Optional nDigits As Integer = 1) As String
    'FPct = ReplaceMinus(Format(d, "0." & RepeatString("0", nDigits) & ChrW(&H200A) & "%")) ' winziger Zwischenraum
    If nDigits >= 1 Then
        FPct = ReplaceMinus(Format(d, "0." & RepeatString("0", nDigits) & "%"))
    Else
        FPct = ReplaceMinus(Format(d, "0%"))
    End If
End Function

Function FPctP(d As Double, Optional nDigits As Integer = 1) As String
    'FPct = ReplaceMinus(Format(d, "0." & RepeatString("0", nDigits) & ChrW(&H200A) & "%")) ' winziger Zwischenraum
    If nDigits >= 1 Then
        FPctP = ReplaceMinus(Format(d, "0." & RepeatString("0", nDigits) & "%p"))
    Else
        FPctP = ReplaceMinus(Format(d, "0%p"))
    End If
End Function

Function FGrowth(d As Double, Optional nDigits As Integer = 1) As String
    Dim pct: pct = FPct(Abs(d), nDigits)
    'FGrowth = ReplaceMinus(Format(d, "+" & ChrW(&H200A) & pct & ";-" & ChrW(&H200A) & pct & ";" & pct)) ' winziger Zwischenraum
    If d > 0 Then
        FGrowth = "+" & pct
    Else
        If d < 0 Then
            FGrowth = ReplaceMinus("-" & pct)
        Else
            FGrowth = pct
        End If
    End If
End Function

Function FGrowthP(d As Double, Optional nDigits As Integer = 1) As String
    Dim growth: growth = FGrowth(d, nDigits)
    FGrowthP = Replace(growth, "%", "") & ChrW(&H200A) & "ppt"
End Function

Function FThousand(d As Double, Optional nDigits As Integer = 0) As String
    If nDigits = 0 Then
        FThousand = ReplaceMinus(Format(d / 1000, "#,##0"))
    Else
        FThousand = ReplaceMinus(Format(d / 1000, "#,##0." & RepeatString("0", nDigits)))
    End If
End Function

Function FMillion(d As Double, Optional nDigits As Integer = 0) As String
    If nDigits = 0 Then
        FMillion = ReplaceMinus(Format(d / 1000000, "#,##0"))
    Else
        FMillion = ReplaceMinus(Format(d / 1000000, "#,##0." & RepeatString("0", nDigits)))
    End If
End Function

Function FThousandK(d As Double, Optional nDigits As Integer = 0) As String
    FThousandK = "€" & ChrW(&H200A) & ReplaceMinus(FThousand(d, nDigits) & ChrW(&H200A) & "k") ' winziger Zwischenraum
End Function

Function FMillionM(d As Double, Optional nDigits As Integer = 0) As String
    FMillionM = "€" & ChrW(&H200A) & ReplaceMinus(FMillion(d, nDigits) & ChrW(&H200A) & "m") ' winziger Zwischenraum
End Function

Function FThousandDeltaK(d As Double, Optional nDigits As Integer = 0) As String
    Dim sThousandK: sThousandK = FThousandK(d, nDigits)
    If sThousandK > 0 Then
        FThousandDeltaK = "+" & sThousandK
    Else
        If sThousandK < 0 Then
            FThousandDeltaK = ReplaceMinus("-" & sThousandK)
        Else
            FThousandDeltaK = sThousandK
        End If
    End If
End Function

Function FMillionDeltaM(d As Double, Optional nDigits As Integer = 0) As String
    Dim sMillionK: sMillionK = FMillionM(d, nDigits)
    If sMillionK > 0 Then
        FMillionDeltaM = "+" & sMillionK
    Else
        If sMillionK < 0 Then
            FMillionDeltaM = ReplaceMinus("-" & sMillionK)
        Else
            FMillionDeltaM = sMillionK
        End If
    End If
End Function

Function FPrice(d As Double, Optional nDigits As Integer = 2) As String
    If nDigits = 0 Then
        ' &H200A: "winziger Zwischenraum"
        ' &H2009: "kleiner Zwischenraum"
        FPrice = ReplaceMinus(Format(d, "€" & ChrW(&H2009) & "0"))
    Else
        FPrice = ReplaceMinus(Format(d, "€" & ChrW(&H2009) & "0." & RepeatString("0", nDigits)))
    End If
End Function

Function FMonth(d As Double, Optional sSuffix As String = "") As String
    If sSuffix = "" Then
        FMonth = Format(d, "MMM" & ChrW(&H200A) & "YY")
    Else
        FMonth = Format(d, "MMM" & ChrW(&H200A) & "YY") & ChrW(&H2009) & sSuffix
    End If
End Function

Function FShortDate(d As Double) As String
    FShortDate = Format(d, "D.M.")
End Function


'*************************************
' Excel SlideText helpers
'*************************************

Function NP(ParamArray v()) As String
    Dim ix, s
    For ix = LBound(v) To UBound(v)
        If v(ix) > 0 Then s = s & "P" Else s = s & "N"
    Next
    NP = s
End Function


Function NP0(rValues As Range, rFormatteds As Range) As String
    Dim rValue As Range, rFormatted As Range
    Dim ix, s
    For ix = 1 To rValues.Cells.Count
        Set rValue = rValues.Cells(ix)
        Set rFormatted = rFormatteds.Cells(ix)
        
        Dim sFormatted As String: sFormatted = Replace(rFormatted.Value2, ChrW(&H2212), "-")
        Dim sNumber As String: sNumber = RegMatch(sFormatted, "[0-9+.,-]+", 0)
        Dim d As Double: d = AsDouble(sNumber)
        If d = 0 Then
            s = s & "0"
        Else
            d = rValue
            If d > 0 Then s = s & "P" Else s = s & "N"
        End If
    Next
    NP0 = s
End Function


Function EN(s As String) As String
    s = Replace(s, ".", "####")
    s = Replace(s, ",", ".")
    s = Replace(s, "####", ",")
    EN = s
End Function


Function ExtractKeywords(ByVal inputString As String) As Collection
    Dim keywords As New Collection
    Dim startPos As Long
    Dim endPos As Long
    Dim keyword As String
    
    startPos = 1
    
    Do
        ' Find next starting $
        startPos = InStr(startPos, inputString, "$")
        If startPos = 0 Then Exit Do
        
        ' Find closing $
        endPos = InStr(startPos + 1, inputString, "$")
        If endPos = 0 Then Exit Do
        
        ' Extract keyword between $ symbols
        keyword = Mid(inputString, startPos + 1, endPos - startPos - 1)
        
        ' Add keyword to collection if not empty
        If Len(Trim(keyword)) > 0 Then
            keywords.Add Trim("$" & keyword & "$")
        End If
        
        ' Move start position after the closing $
        startPos = endPos + 1
    Loop

    Set ExtractKeywords = keywords
End Function


Sub SimulateSlideText()
    Debug.Assert NamedRangeExists(R_SLIDE_TEXTS, ActiveSheet)
    Dim rSlideTexts As Range: Set rSlideTexts = Range(R_SLIDE_TEXTS)
    Dim rSelection As Range: Set rSelection = Intersect(Selection.EntireRow, rSlideTexts)
    Debug.Assert Not rSelection Is Nothing

    Dim iUseCol, iPatternCol, iGroupCol, iReplaceCol
    If rSlideTexts.Columns.Count = 3 Then
        iUseCol = -1: iPatternCol = 1: iGroupCol = 2: iReplaceCol = 3
    Else
        iUseCol = 1: iPatternCol = 2: iGroupCol = 3: iReplaceCol = 4
    End If
    
    Dim sPattern As String: sPattern = rSelection.Cells(iPatternCol)
    Dim rFirstPattern As Range
    Dim ixFirstRow, ixLastRow
    For ixFirstRow = 1 To rSlideTexts.Rows.Count
        If rSlideTexts.Cells(ixFirstRow, iPatternCol) = sPattern Then Exit For
    Next
    For ixLastRow = rSlideTexts.Rows.Count To 1 Step -1
        If rSlideTexts.Cells(ixLastRow, iPatternCol) = sPattern Then Exit For
    Next
    
    Dim cKeywords As Collection: Set cKeywords = ExtractKeywords(sPattern)
    
    Dim nGroups: nGroups = ixLastRow - ixFirstRow + 1
    Dim ixGroup
    For ixGroup = 1 To Min(nGroups, cKeywords.Count)
        Dim sKeyword As String: sKeyword = cKeywords(ixGroup)
        Dim ixFound As Integer: ixFound = InStr(sPattern, sKeyword)
        Debug.Assert ixFound > 0
        Dim sGroup As String: sGroup = rSlideTexts.Cells(ixFirstRow + ixGroup - 1, iReplaceCol)
        sPattern = Left(sPattern, ixFound - 1) & sGroup & Mid(sPattern, ixFound + Len(sKeyword))
    Next
    MsgBox sPattern
End Sub



'>> RangeGetters

' refactor

'*************************************
' offset
'*************************************

Function FastOffset(r As Range, iRowOffs As Integer, iColOffs As Integer, Optional nRows As Integer, Optional nCols As Integer) As Range
    ' BEREICH.VERSCHIEBEN is volatile, this function isn't
    ' NOTE the problem that it is usually necessary to combine this w/ a Link2 or Link3

    If nRows = 0 Then nRows = 1
    If nCols = 0 Then nCols = 1
    If nRows = 1 And nCols = 1 Then
        Set FastOffset = r.Offset(iRowOffs, iColOffs)
    Else
        Dim rTopLeft As Range: Set rTopLeft = r.Offset(iRowOffs, iColOffs)
        Dim rBottomRight As Range: Set rBottomRight = rTopLeft.Offset(nRows - 1, nCols - 1)
        Set FastOffset = rTopLeft.Worksheet.Range(rTopLeft, rBottomRight)
    End If
End Function


Function CompactVector(rVector As Range) As Range
    Dim rFirst As Range: Set rFirst = FirstInRange(rVector)
    Dim rLast As Range: Set rLast = LastInRange(rVector)
    If IsEmpty(rFirst) Then Set rFirst = rFirst.End(xlToRight)
    If IsEmpty(rLast) Then Set rLast = rLast.End(xlToLeft)
    Set CompactVector = rVector.Worksheet.Range(rFirst, rLast)
End Function


'*************************************
' single subcell
'*************************************

Function FirstInRange(r As Range) As Range
    If Not r Is Nothing Then Set FirstInRange = r.Cells(1)
End Function

Function LastInRange(r As Range) As Range
    If Not r Is Nothing Then Set LastInRange = r.Cells(r.Cells.Count)
End Function


'*************************************
' new last empty row/col
'*************************************

Function RangeOrientationOld(r As Range) As Integer
    Dim rLast As Range: Set rLast = r.SpecialCells(xlCellTypeLastCell)
    If rLast.Column = r.Column And rLast.Row = r.Row Then
            
        ' ... so check if it might be a whole column or row
        ' this is amazingly fast, so no worries :)
        Dim rOffset As Range
        On Error Resume Next
        
        ' first check if it is possible to select the next column
        Set rOffset = r.Offset(0, 1)
        If rOffset Is Nothing Then
            ' it is not possible to select the next column, because there is no next column b/o we passed all columns in r
            RangeOrientationOld = xlOrientationRow
        Else
            ' it is possible to select the column to the right of r, but we still don't know if it is a complete worksheet column or actually a single cell
            ' => check if we can move to next row
            Set rOffset = Nothing: Set rOffset = r.Offset(1, 0)
            If rOffset Is Nothing Then
                ' we already passed all rows in r, so Offset cannot go to the next row because there isn't any
                RangeOrientationOld = xlOrientationColumn
            Else
                ' it is actually a single cell
                RangeOrientationOld = xlOrientationCell
            End If
        End If
    Else
        If rLast.Column = r.Column Then
            RangeOrientationOld = xlOrientationColumn
        Else
            If rLast.Row = r.Row Then
                RangeOrientationOld = xlOrientationRow
            Else
                ' both last row and column cell are different from passed row
                RangeOrientationOld = xlOrientationMatrix
            End If
        End If
    End If
End Function


Function RangeOrientation(r As Range) As Integer
    ' SpecialCells(A:A) returns A:A, so we would errornously get "single cell" as result ...
    ' 13.02.18 unklar unter welchen Bedingungen SpecialCells *nicht* funktioniert, möglicherweise bei bestimmten UDFs, lässt sich aber nicht reproduzieren
    '    Problem mit SpecialCells trat mit CreateHeaderDictionary auf
    '    SpecialCells funktioniert auch nicht richtig, wenn die Zellen (wie bei import sheets z.B.) formatiert sind.
    ' (see https://bettersolutions.com/excel/cells-ranges/vba-special-cells.htm for BUG for more than 16.385 rows)
    ' => we use Offset() to induce error which is then used as indication for orientation
    
    If r Is Nothing Then
        RangeOrientation = -1
    Else
        If r.Cells.Count = 1 Then
            
            ' counting cells is fast, as opposed to counting rows/columns (see below)
            RangeOrientation = xlOrientationCell
        Else
            ' Rows.Count and/or Columns.Count massively slows down this method, so we can't use them
            'If r.Rows.Count = 0 Then
            '    RangeOrientation = xlOrientationColumn
            'Else
            '    If r.Columns.Count > 0 Then
            '        RangeOrientation = xlOrientationMatrix
            '    Else
            '        RangeOrientation = xlOrientationRow
            '    End If
            'End If
            'Exit Function
    
            ' first check if it is possible to select the next column
            On Error Resume Next
            Dim rOffsetCol As Range: Set rOffsetCol = r.Offset(0, 1)
            If rOffsetCol Is Nothing Then
                ' it is not possible to select the next column, because there is no next column b/o we passed all columns in r
                RangeOrientation = xlOrientationRow
            Else
                ' it is possible to select the column to the right of r, but we still don't know if it is a complete worksheet column or actually a single cell
                ' => check if we can move to next row
                Dim rOffsetRow As Range: Set rOffsetRow = r.Offset(1, 0)
                If rOffsetRow Is Nothing Then
                    ' we already passed all rows in r, so Offset cannot go to the next row because there isn't any
                    RangeOrientation = xlOrientationColumn
                Else
                    ' offset in both directions was successful
                    ' it can be anything
                
                    If Intersect(rOffsetCol, r) Is Nothing Then
                        ' offset to the next col takes us out of the passed range
                        
                        ' already handled in ((FITVPZW)) above
                        'If Intersect(rOffsetRow, r) Is Nothing Then
                        '    RangeOrientation = xlOrientationCell
                        'Else
                            RangeOrientation = xlOrientationColumn
                        'End If
                    Else
                        ' after offset to the next col we are still in the passed range
                    
                        If Intersect(rOffsetRow, r) Is Nothing Then
                            RangeOrientation = xlOrientationRow
                        Else
                            RangeOrientation = xlOrientationMatrix
                        End If
                    End If
                
                End If
            End If
        End If
    End If
End Function


Function LastNonEmptyInRange(r As Range) As Range
    ' idea: do not use Cells.Count on passed r (but see matrix)
    
    If r Is Nothing Then
        Set LastNonEmptyInRange = Nothing
    Else
        Select Case RangeOrientation(r)
            Case xlOrientationColumn
                Set LastNonEmptyInRange = r.Worksheet.Cells(EXCEL_LAST_ROW, r.Column).End(xlUp) 'r.Worksheet.Rows.Count
            Case xlOrientationRow
                Set LastNonEmptyInRange = r.Worksheet.Cells(r.Row, EXCEL_LAST_COL).End(xlLeft) ' r.Worksheet.Columns.Count
                
            Case xlOrientationCell
                Set LastNonEmptyInRange = r
                
            Case xlOrientationMatrix
                Dim ixRow As Long
                
                ' we use Count here, which is slow, but we usually use LastNonEmptyInRange() for (row/col) vectors anyways, so let's take the minimal performance hit here
                For ixRow = r.Rows.Count To 1 Step -1
                    Dim rRow As Range: Set rRow = r.Rows(ixRow)
                    If WorksheetFunction.CountA(rRow) > 0 Then
                        Set LastNonEmptyInRange = LastNonEmptyInRange(rRow)
                        Exit Function
                    End If
                Next
                Set LastNonEmptyInRange = Nothing
        End Select
    End If
End Function


Function FirstNonEmptyInRange(r As Range) As Range
    ' 13.02.18 not thoroughly tested
    
    If r Is Nothing Then
        Set FirstNonEmptyInRange = Nothing
    Else
        Dim rFirst As Range: Set rFirst = r.Cells(1)
        If Not IsEmpty(rFirst) Then
            Set FirstNonEmptyInRange = rFirst
        Else
            Select Case RangeOrientation(r)
                Case xlOrientationColumn
                    Set FirstNonEmptyInRange = rFirst.End(xlDown)
                Case xlOrientationRow
                    Set FirstNonEmptyInRange = rFirst.End(xlToLeft)
                Case xlOrientationCell
                    Set FirstNonEmptyInRange = r
                Case xlOrientationMatrix
                    Dim rRow As Range
                    For Each rRow In r.Rows
                        ' NOTE: CountA is internally obviously also a .Cells.Count, i.e. is rather slow
                        If WorksheetFunction.CountA(rRow) > 0 Then
                            Set FirstNonEmptyInRange = FirstNonEmptyInRange(rRow)
                            Exit Function
                        End If
                    Next
                    Set FirstNonEmptyInRange = Nothing
            End Select
        End If
    End If
End Function



'*************************************
' indirect NonZero
'*************************************

Function IndirectFirstNonZero(rLookup As Range, rReturn As Range)
    Dim r As Range
    Dim ix As Integer
    For ix = 1 To rLookup.Count
        Set r = rLookup.Cells(ix)
        If Not IsEmpty(r) Then
            If IsNumeric(r) Then
                If r <> 0 Then
                    Debug.Assert rReturn.Cells.Count >= ix
                    Set IndirectFirstNonZero = rReturn.Cells(ix)
                    Exit Function
                End If
            End If
        End If
    Next
    IndirectFirstNonZero = ErrNA()
End Function


Function IndirectLastNonZero(rLookup As Range, rReturn As Range)
    Dim r As Range
    Dim ix As Integer
    For ix = rLookup.Count To 1 Step -1
        Set r = rLookup.Cells(ix)
        If Not IsEmpty(r) Then
            If IsNumeric(r) Then
                If r <> 0 Then
                    Debug.Assert rReturn.Cells.Count >= ix
                    Set IndirectLastNonZero = rReturn.Cells(ix)
                    Exit Function
                End If
            End If
        End If
    Next
    IndirectLastNonZero = ErrNA()
End Function


'*************************************
' check range for empty and return idx
'*************************************

Function FirstNumberIndex(rNumbers As Range) As Integer
    Dim rNumber As Range
    Dim ixNumber As Integer
    For ixNumber = 1 To rNumbers.Cells.Count
        Set rNumber = rNumbers.Cells(ixNumber)
        If Not IsEmpty(rNumber) And IsNumeric(rNumber) Then
            FirstNumberIndex = ixNumber
            Exit Function
        End If
    Next
    FirstNumberIndex = -1
End Function

Function LastNumberIndex(rNumbers As Range) As Integer
    Dim rNumber As Range
    Dim ixNumber As Integer
    For ixNumber = rNumbers.Cells.Count To 1 Step -1
        Set rNumber = rNumbers.Cells(ixNumber)
        If Not IsEmpty(rNumber) And IsNumeric(rNumber) Then
            LastNumberIndex = ixNumber
            Exit Function
        End If
    Next
    LastNumberIndex = -1
End Function


'*************************************
' old last/first
'*************************************

Function LastInColumn(rColumn As Range) As Range
    Dim rSourceTop As Range: Set rSourceTop = rColumn.Cells(1)
    Dim rLastInColumn As Range: Set rLastInColumn = rColumn.Cells(rColumn.Cells.Count)
    Dim rSourceEndUp As Range: Set rSourceEndUp = rLastInColumn.End(xlUp)
    Set LastInColumn = rColumn.Worksheet.Range(rSourceTop, rSourceEndUp)
End Function

Function LastInColumn2(rColumn As Range) As Range
    Dim rSourceTop As Range: Set rSourceTop = rColumn.Cells(1)
    Dim rLastInColumn As Range: Set rLastInColumn = rColumn.Cells(rColumn.Cells.Count)
    Dim rSourceEndUp As Range: Set rSourceEndUp = rLastInColumn.End(xlUp)
    Set LastInColumn2 = rSourceEndUp
End Function

Function IsFirstInRange(r As Range, rMaybeFirst As Range) As Boolean
    IsFirstInRange = FirstInRange(r).Address = rMaybeFirst.Address
End Function


Function IsLastInRange(r As Range, rMaybeLast As Range) As Boolean
    IsLastInRange = LastInRange(r).Address = rMaybeLast.Address
End Function


Function LastNonZeroInRange(r As Range) As Range
    Dim ix As Integer
    For ix = r.Cells.Count To 1 Step -1
        If r.Cells(ix) <> 0 Then
            Set LastNonZeroInRange = r.Cells(ix)
            Exit Function
        End If
    Next
    Set LastNonZeroInRange = Nothing
End Function


Function SafeEndToRight(rLeft As Range) As Range
    Dim r As Range: Set r = rLeft
    Dim rLast As Range
    Do
        Set rLast = r
        Set r = r.Offset(0, 1)
    Loop While Not IsEmpty(r)
    Set SafeEndToRight = rLast
End Function


'*************************************
' subranges
'*************************************

Private Function LastNonEmptyColumn(r As Range) As Range
    Debug.Assert r.Rows.Count = 1
    Dim rRight As Range: Set rRight = r.Cells(r.Columns.Count)
    If IsEmpty(rRight) Then
        Dim rLastFromRight As Range: Set rLastFromRight = rRight.End(xlToLeft)
        If rLastFromRight.Column = 1 And r.Column <> 1 Then
            Set LastNonEmptyColumn = r
        Else
            Set LastNonEmptyColumn = rLastFromRight
        End If
    Else
        Set LastNonEmptyColumn = rRight
    End If
End Function


Function GetSubrange(r As Range, iRowOffset As Integer, iColumnOffset As Integer, nRows As Integer, nColumns As Integer) As Range
    Dim rTopLeft As Range: Set rTopLeft = r.Cells(1, 1).Offset(iRowOffset, iColumnOffset)
    Dim rBottomRight As Range: Set rBottomRight = rTopLeft.Offset(nRows - 1, nColumns - 1)
    Set GetSubrange = r.Worksheet.Range(rTopLeft, rBottomRight)
End Function

Function GetDataRangeFromFirstColumnAndFirstRow(rTopLeft As Range) As Range
    Dim rTopRight As Range: Set rTopRight = rTopLeft.End(xlToRight)
    Dim rBottomLeft As Range: Set rBottomLeft = rTopLeft.End(xlDown)
    Dim rBottomRight As Range: Set rBottomRight = Intersect(rTopRight.EntireColumn, rBottomLeft.EntireRow)
    Set GetDataRangeFromFirstColumnAndFirstRow = rTopLeft.Worksheet.Range(rTopLeft, rBottomRight)
End Function


Function GetDataRangeEndDownEndRight(rTopLeft As Range) As Range
    Dim rBottomRight As Range: Set rBottomRight = rTopLeft.End(xlDown).End(xlToRight)
    Set GetDataRangeEndDownEndRight = rTopLeft.Worksheet.Range(rTopLeft, rBottomRight)
End Function


Function GetOuterTable(rInsideRange As Range) As Range
    Set GetOuterTable = rInsideRange.Worksheet.Range(rInsideRange, rInsideRange.SpecialCells(xlLastCell))
End Function


Function DetermineColumnRows(rHeader As Range, sColumn As String) As Long
    Dim ixColumn As Long: ixColumn = match(sColumn, rHeader)
    Debug.Assert ixColumn > 0
    Dim rSourceTop As Range: Set rSourceTop = rHeader.Cells(ixColumn).Offset(1, 0)
    Dim rSourceEndDown As Range: Set rSourceEndDown = rSourceTop.End(xlDown)
    DetermineColumnRows = rSourceEndDown.Row - rSourceTop.Row + 1
End Function


Function DetermineColumnRows2(rHeader As Range, sColumn As String) As Long
    Dim ixColumn As Long: ixColumn = match(sColumn, rHeader)
    Debug.Assert ixColumn > 0
    Dim rSourceTop As Range: Set rSourceTop = rHeader.Cells(ixColumn).Offset(1, 0)
    Dim rSourceEndDown As Range: Set rSourceEndDown = rSourceTop.EntireColumn.SpecialCells(xlCellTypeLastCell)
    If rSourceEndDown.Row = EXCEL_LAST_ROW Then
        DetermineColumnRows2 = -1
    Else
        DetermineColumnRows2 = rSourceEndDown.Row - rSourceTop.Row + 1
    End If
End Function


Function DetermineColumnRows3(rHeader As Range, sColumn As String) As Long
    Dim ixColumn As Long: ixColumn = match(sColumn, rHeader)
    Debug.Assert ixColumn > 0
    Dim rSourceTop As Range: Set rSourceTop = rHeader.Cells(ixColumn).Offset(1, 0)
    With rSourceTop.EntireColumn
        Dim rLastInColumn As Range: Set rLastInColumn = .Cells(.Cells.Count)
    End With
    Dim rSourceEndUp As Range: Set rSourceEndUp = rLastInColumn.End(xlUp)
    DetermineColumnRows3 = rSourceEndUp.Row - rSourceTop.Row + 1
End Function


Function GetColumnRangeFromAnchor(rAnchor As Range) As Range
    Set GetColumnRangeFromAnchor = rAnchor.Worksheet.Range(rAnchor, rAnchor.End(xlDown))
End Function


Function GetColumnRangeFromAnchorVolatile(rAnchor As Range) As Range
    Application.Volatile
    Set GetColumnRangeFromAnchorVolatile = rAnchor.Worksheet.Range(rAnchor, rAnchor.End(xlDown))
End Function


Function GetDataRange(rDataAnchor As Range, nRows, Optional nCols As Integer) As Range
    If nCols > 0 Then
        Set GetDataRange = rDataAnchor.Worksheet.Range(rDataAnchor, rDataAnchor.Offset(nRows - 1, nCols - 1))
    Else
        Set GetDataRange = rDataAnchor.Worksheet.Range(rDataAnchor, rDataAnchor.End(xlToRight).Offset(nRows - 1, 0))
    End If
End Function


Function GetDataColumn(rHeader As Range, sColHeader As String, Optional nRows As Long) As Range
    Dim ixCol As Integer: ixCol = match(sColHeader, rHeader)
    
    ' return Nothing if unknown header
    If ixCol = 0 Then Exit Function
        
    
    Dim rFirst As Range: Set rFirst = rHeader.Cells(ixCol).Offset(1)
    
    Dim rLast As Range
    If nRows = 0 Then
        ' dangerous, because the column might have empty cells
        Set rLast = rFirst.End(xlDown)
    Else
        Set rLast = rFirst.Offset(nRows - 1)
    End If

    Set GetDataColumn = rHeader.Worksheet.Range(rFirst, rLast)
    
End Function


Function CountRows(rHeader As Range, sColHeader As String) As Long
    Dim ixCol As Integer: ixCol = match(sColHeader, rHeader)
    If ixCol > 0 Then
    
        Dim rFirst As Range: Set rFirst = rHeader.Cells(ixCol).Offset(1)
        Dim rLast As Range: Set rLast = rFirst.End(xlDown)
    
        CountRows = rLast.Row - rFirst.Row + 1
    
    Else
    
        ' design error
        Debug.Assert False
    
    End If
End Function


'*************************************
' cropping a named range
'*************************************

Function CropColumnRangeOld(ws As Worksheet, rSafeRaw As Range) As Range
    Set CropColumnRangeOld = ws.Range(FirstInRange(rSafeRaw).Offset(1), LastInRange(rSafeRaw).Offset(-1))
End Function

Function CropRowRangeOld(ws As Worksheet, rSafeRaw As Range) As Range
    Set CropRowRangeOld = ws.Range(FirstInRange(rSafeRaw).Offset(0, 1), LastInRange(rSafeRaw).Offset(0, -1))
End Function

Function CropRange(ws As Worksheet, rSafeRaw As Range) As Range
    Dim ix
    Dim rFirst As Range: Set rFirst = Nothing
    For ix = 1 To rSafeRaw.Cells.Count
        If Not IsEmpty(rSafeRaw.Cells(ix)) Then
            Set rFirst = rSafeRaw.Cells(ix)
            Exit For
        End If
    Next
    If rFirst Is Nothing Then Exit Function
    
    
    Dim rLast As Range: Set rLast = Nothing
    For ix = rSafeRaw.Cells.Count To 1 Step -1
        If Not IsEmpty(rSafeRaw.Cells(ix)) Then
            Set rLast = rSafeRaw.Cells(ix)
            Exit For
        End If
    Next
    If rLast Is Nothing Then Exit Function
    Set CropRange = ws.Range(rFirst, rLast)
End Function

Function CropColumnRange(ws As Worksheet, rSafeRaw As Range) As Range
    Set CropColumnRange = CropRange(ws, rSafeRaw)
End Function

Function CropRowRange(ws As Worksheet, rSafeRaw As Range) As Range
    Set CropRowRange = CropRange(ws, rSafeRaw)
End Function



'*************************************
' accessing column information
'*************************************

Function CollectConstantsAndFormulas(r As Range) As Range
    ' IMPORTANT: .SpecialCells only works in subs, not functions!
    ' => use CompactVector instead
    ' set ((RIMCNHB)) below
    
    On Error Resume Next
    Dim rConstants As Range: Set rConstants = r.Cells.SpecialCells(xlCellTypeConstants)
    Dim rFormulas As Range: Set rFormulas = r.Cells.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    Dim rUnion As Range
    If rConstants Is Nothing And rFormulas Is Nothing Then
        ' return nothing
    Else
        If rConstants Is Nothing Then
            ' rFormulas <> nothing
            Set CollectConstantsAndFormulas = rFormulas
        Else
            If rFormulas Is Nothing Then
                Set CollectConstantsAndFormulas = rConstants
            Else
                Set CollectConstantsAndFormulas = Union(rConstants, rFormulas)
            End If
        End If
    End If
End Function


Function CreateHeaderDictionary(rHeader As Range) As Dictionary
    ' Accessing the columns via the string header needs a way to resolve the string to a column number.
    ' => load the column header strings in a dictionary
    ' This is MUCH faster than iterating over the columns to find the correct column.

    ' set ((RIMCNHB)) above
    Dim rHeaderCells As Range
    'Set rHeaderCells = CollectConstantsAndFormulas(rHeader)
    Set rHeaderCells = CompactVector(rHeader.Cells)

    Dim dict As New Dictionary
    Dim rColHeader As Range
    For Each rColHeader In rHeaderCells
        Dim sColHeader As String: sColHeader = Replace(Trim(rColHeader.Value2), Chr(10), " ")
        If sColHeader = "" Then
            ' skip empty columns (more precise: columns which don't have a header)
        Else
            If dict.Exists(sColHeader) Then
                ' skip this col, already occured left (i.e. take the first)
            Else
                dict.Add sColHeader, rColHeader.Column
            End If
        End If
    Next
    Set CreateHeaderDictionary = dict
End Function


Function GetColIndex(dictHeader As Dictionary, sColHeader As String) As Integer
    GetColIndex = dictHeader(sColHeader)
End Function


Function GetColValue(r As Range, dictHeader As Dictionary, sColHeader As String) As Variant
    Dim rRow As Range: Set rRow = r.EntireRow
    Dim rCell As Range: Set rCell = rRow.Cells(GetColIndex(dictHeader, sColHeader))
    
    ' intentionally return an error, caller should not auto-cast into string
    GetColValue = rRow.Cells(GetColIndex(dictHeader, sColHeader))
End Function


Sub SetColValue(r As Range, dictHeader As Dictionary, sColHeader As String, value As Variant)
    Dim rRow As Range: Set rRow = r.EntireRow
    Dim rCell As Range: Set rCell = rRow.Cells(GetColIndex(dictHeader, sColHeader))
    rCell.Value2 = value
End Sub


'*************************************
' range objects
'*************************************

Function CreateRange(sName As String, r As Range) As String
    If m_RangeObjects Is Nothing Then Set m_RangeObjects = New Dictionary
    sName = UCase(Trim(sName))
    If m_RangeObjects.Exists(sName) Then m_RangeObjects.Remove sName
    m_RangeObjects.Add sName, r
    CreateRange = sName
End Function

Function GetRange(sName As String) As Range
    sName = UCase(Trim(sName))
    If m_RangeObjects.Exists(sName) Then
        Set GetRange = m_RangeObjects(sName)
    Else
        GetRange = "#unknown"
    End If
End Function


'*************************************
' diffference
'*************************************

' Intersect, Union implemented, but not difference

Function Difference(r1 As Range, r2 As Range) As Range
    ' https://www.mrexcel.com/board/threads/what-is-the-best-way-to-do-an-opposite-intersect.613349/
    Dim s As String
    Dim ws As Worksheet
     
    If Not (r1.Parent Is r2.Parent) Then Exit Function
    On Error Resume Next
     
    Set ws = Worksheets.Add
    ws.Range(r1.Address) = 0
    ws.Range(r2.Address).Clear
    s = ws.Range(r1.Address).SpecialCells(xlCellTypeConstants).Address
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    If s <> "" Then Set Difference = r1.Parent.Range(s)
End Function



'>> RegExpr

' http://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops

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


'*************************************
' matching
'*************************************

' main function

Function CheckMissing(s As String, Optional ixGroup) As Boolean
    CheckMissing = IsMissing(ixGroup)
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
    
    
Function RegReplace(vTarget, sPattern As String, sReplace As String, Optional bCaseSensitive As Boolean, Optional bMultiLine As Boolean, Optional bGlobal As Boolean = True) As Variant
    Dim sTarget As String: sTarget = vTarget
    Dim RE As RegExp: Set RE = SafeCreateRegExp(sPattern, bCaseSensitive, bMultiLine, bGlobal)
    RegReplace = RE.Replace(sTarget, sReplace)
End Function
    
    

'*************************************
' helpers
'*************************************

Sub RemoveRegexChars(ByRef sfnCompact As String)
    sfnCompact = Replace(sfnCompact, "(", "\(")
    sfnCompact = Replace(sfnCompact, ")", "\)")
    sfnCompact = Replace(sfnCompact, "[", "\[")
    sfnCompact = Replace(sfnCompact, "]", "\]")
    sfnCompact = Replace(sfnCompact, "+", "\+")
End Sub



'>> Reporting

Function SumUp(rKonten As Range, rValues As Range) As Double
    ' all 2 passed ranges must have the same range and start in the same row
    ' this is usually the case anyways, just checking here
    Dim n As Integer: n = rKonten.Rows.Count
    Debug.Assert rValues.Rows.Count = n
    Debug.Assert rKonten.Rows(1).Row = rValues.Rows(1).Row
    Debug.Assert rKonten.Columns.Count = 1

    Dim rCaller As Range: Set rCaller = Application.Caller
    
    Dim ixRelativeCallerRow As Integer: ixRelativeCallerRow = rCaller.Row - rKonten.Rows(1).Row + 1
    
    Dim ixRelativeValueRow As Integer: ixRelativeValueRow = ixRelativeCallerRow - 1

    Dim dSum As Double
    Do While ixRelativeValueRow > 0
        
        Dim bIsKonto As Boolean: bIsKonto = rKonten.Rows(ixRelativeValueRow)
        If Not bIsKonto Then
        
            ' do not add this line
            ' it is probably a sum or remark or something else
        
        Else
        
            ' it is a Konto, so sum the value
            dSum = dSum + rValues.Rows(ixRelativeValueRow)
        
        End If
        
        ixRelativeValueRow = ixRelativeValueRow - 1
        
    Loop
    
    SumUp = dSum
    
End Function


Function SumDown(rDescriptions As Range, rKonten As Range, rValues As Range) As Double
    ' all 3 passed ranges must have the same range and start in the same row
    ' this is usually the case anyways, just checking here
    Dim n As Integer: n = rDescriptions.Rows.Count
    Debug.Assert rKonten.Rows.Count = n
    Debug.Assert rValues.Rows.Count = n
    Debug.Assert rDescriptions.Rows(1).Row = rKonten.Rows(1).Row
    Debug.Assert rDescriptions.Rows(1).Row = rValues.Rows(1).Row
    
    Dim rCaller As Range: Set rCaller = Application.Caller
    
    rCaller.Worksheet.Activate
    
    Dim ixRelativeCallerRow As Integer: ixRelativeCallerRow = rCaller.Row - rDescriptions.Rows(1).Row + 1
    Dim iCallerIndent As Integer: iCallerIndent = rDescriptions.Rows(ixRelativeCallerRow).IndentLevel
    
    Dim ixRelativeValueRow As Integer: ixRelativeValueRow = ixRelativeCallerRow + 1
    Dim dSum As Double
    Do While ixRelativeValueRow <= n
    
        Dim rDescription As Range: Set rDescription = rDescriptions.Rows(ixRelativeValueRow)
        
        If rDescription.IndentLevel <= iCallerIndent Then
            
            Exit Do
            
        Else
            Dim bIsKonto As Boolean: bIsKonto = rKonten.Rows(ixRelativeValueRow)
            If Not bIsKonto Then
            
                ' do not add this line
                ' it is probably a sum or remark or something else
            
            Else
            
                ' it is a Konto, so sum the value
                dSum = dSum + rValues.Rows(ixRelativeValueRow)
            
            End If
            
            ixRelativeValueRow = ixRelativeValueRow + 1
            
        End If
        
    Loop
    
    SumDown = dSum
End Function


Function SumDown2(rDescriptions As Range, rKonten As Range, rIndents As Range, rValues As Range) As Double
    ' 02.01.2018
    ' unclear why I ever introduced SumDown2()
    ' This function is exactly like SumDown() *BUT* has passed additional rIndents.
    ' The reason might have been that in the beginning we used a volatile GetIndent() ... but this doesn't work as intendend...
    ' ... because Range.IndentLevel seems to not work in certain situations (if I remember correctly: on startup, or on remove some sheet)
    
    Stop
    ' ... try to use SumDown()
    
    ' all 3 passed ranges must have the same range and start in the same row
    ' this is usually the case anyways, just checking here
    Dim n As Integer: n = rDescriptions.Rows.Count
    Debug.Assert rKonten.Rows.Count = n
    Debug.Assert rValues.Rows.Count = n
    Debug.Assert rDescriptions.Rows(1).Row = rKonten.Rows(1).Row
    Debug.Assert rDescriptions.Rows(1).Row = rValues.Rows(1).Row
    
    Dim rCaller As Range: Set rCaller = Application.Caller
    
    rCaller.Worksheet.Activate
    
    Dim ixRelativeCallerRow As Integer: ixRelativeCallerRow = rCaller.Row - rDescriptions.Rows(1).Row + 1
    'Dim iCallerIndent As Integer: iCallerIndent = rDescriptions.Rows(ixRelativeCallerRow).IndentLevel
    Dim iCallerIndent As Integer: iCallerIndent = rIndents.Rows(ixRelativeCallerRow).Value2
    
    Dim ixRelativeValueRow As Integer: ixRelativeValueRow = ixRelativeCallerRow + 1
    Dim dSum As Double
    Do While ixRelativeValueRow <= n
    
        Dim rDescription As Range: Set rDescription = rDescriptions.Rows(ixRelativeValueRow)
        
        If rDescription.IndentLevel <= iCallerIndent Then
            
            Exit Do
            
        Else
            Dim bIsKonto As Boolean: bIsKonto = rKonten.Rows(ixRelativeValueRow)
            If Not bIsKonto Then
            
                ' do not add this line
                ' it is probably a sum or remark or something else
            
            Else
            
                ' it is a Konto, so sum the value
                dSum = dSum + rValues.Rows(ixRelativeValueRow)
            
            End If
            
            ixRelativeValueRow = ixRelativeValueRow + 1
            
        End If
        
    Loop
    
    SumDown2 = dSum
    
End Function



'>> SheetJump

'*************************************
' stuff from Soli (row height change)
'*************************************

Sub Button_HideDisplayHeadingsOnSections()
    HideDisplayHeadingsOnSections
End Sub

Sub Button_HideDisplayHeaders()
    HideDisplayHeadings
End Sub

Sub Button_FreezeAllPanes()
    FreezeAllPanes
End Sub

Sub HideDisplayHeadingsOnSections()
    Dim wsOld: Set wsOld = ActiveSheet
    Dim ws As Worksheet
    For Each ws In Sheets
        If IsSectionSheet(ws) Then
            ws.Activate
            ActiveWindow.DisplayHeadings = False
        End If
    Next
    wsOld.Activate
End Sub


Sub HideDisplayHeadings()
    SetSilentApplicationState
    Dim old: Set old = ActiveSheet
    Dim ws As Worksheet
    For Each ws In Worksheets
        If IsSectionSheet(ws) Then
            ws.Activate
            ActiveWindow.DisplayHeadings = False
        End If
    Next
    old.Activate
    RevertApplicationState
End Sub


Sub FreezeAllPanes()
    Dim old: Set old = ActiveSheet
    Dim ws As Worksheet
    For Each ws In Worksheets
        If NamedRangeExists("_", ws) Then
            ws.Activate
            ws.Range("_").Select
            ActiveWindow.FreezePanes = True
        End If
    Next
    old.Activate
End Sub


Sub DefineAnchorsForFrozenPanes()
    ' https://www.mrexcel.com/board/threads/vba-how-do-i-figure-out-what-cell-at-which-panes-are-frozen-with-freezepanes.1129717/
    Dim wsOld: Set wsOld = ActiveSheet
    Dim ws As Worksheet
    For Each ws In Sheets
        If ActiveWindow.FreezePanes Then
            
        End If
    Next
    wsOld.Activate
End Sub


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


'*************************************
' sections (smaller steps) - ^j, ^k
'*************************************

Sub ColorAsSectionSheet()
    ActiveSheet.Tab.Color = RGB(43, 22, 75)
End Sub


Sub FormatArrowSectionSheets()
    SetSilentApplicationState
    Dim wsOld: Set wsOld = ActiveSheet
    Dim ws As Worksheet
    For Each ws In Worksheets
        If IsSectionSheet(ws, True) Then
            ws.Activate
            ActiveWindow.DisplayHeadings = False
            With ws.Range("G25")
                .Value2 = ws.Name
                .Font.Size = 12
                .Font.Bold = True
                .RowHeight = 15.75
            End With
            
        End If
    Next
    wsOld.Activate
    RevertApplicationState
End Sub


Sub HideHeadingsOnAllSectionSheets()
    Dim ws As Worksheet
    For Each ws In Worksheets
        If IsSectionSheet(ws) Then
            ws.Activate
            ActiveWindow.DisplayHeadings = False
        End If
    Next
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


Sub GotoNextPart()
    Dim ixSheet As Integer
    
    ' find next part sheet; +1: we might start looking while being on a part sheet
    For ixSheet = ActiveSheet.index + 1 To Worksheets.Count
        Dim ws As Worksheet: Set ws = Worksheets(ixSheet)
        
        ' check if Registerfarbe indicates that it is a part sheet
        If IsPartSheet(ws) Then
            ws.Activate
            Exit Sub
        End If
        
    Next ixSheet
End Sub


Sub GotoPrevPart()
    Dim ixSheet As Integer
    
    ' find next section sheet; -1: see GotoNextPart
    For ixSheet = ActiveSheet.index - 1 To 1 Step -1
        Dim ws As Worksheet: Set ws = Worksheets(ixSheet)
        
        ' check if Registerfarbe indicates that it is a part sheet
        If IsPartSheet(ws) Then
            ws.Activate
            Exit Sub
        End If
        
    Next ixSheet
End Sub


'*************************************
' circles ^q, ^w, ^e
'*************************************

Function CircleSheetsRange() As Range
    Set CircleSheetsRange = Worksheets("Sheets").Range("$A:$A")
End Function

Sub RemoveSheetFromCircle(sSheetName As String, Optional sMessage As String)
    Dim rCircle As Range: Set rCircle = CircleSheetsRange()
    Dim ixToRemove As Integer: ixToRemove = match(sSheetName, rCircle)
    Dim bSheetInCircle As Boolean: bSheetInCircle = ixToRemove > 0
    If bSheetInCircle Then
        rCircle.Cells(ixToRemove).EntireRow.Delete
        If sMessage = "" Then sMessage = sSheetName & " removed from circle"
        Application.StatusBar = sMessage
    End If
End Sub

Sub RemoveActiveSheetFromCircle()
    RemoveSheetFromCircle ActiveSheet.Name
End Sub

Sub AddActiveSheetToCircle()
    Dim rCircle As Range: Set rCircle = CircleSheetsRange()
    Dim nSheetsInCircle As Integer: nSheetsInCircle = Application.WorksheetFunction.CountA(rCircle)
    rCircle.Cells(nSheetsInCircle + 1) = ActiveSheet.Name
    Application.StatusBar = ActiveSheet.Name & " added to circle"
End Sub


Sub ToggleActiveSheetInCircle()
    Dim rCircle As Range: Set rCircle = CircleSheetsRange()
    Dim ixToRemove As Integer: ixToRemove = match(ActiveSheet.Name, rCircle)
    Dim bSheetAlreadyInCircle: bSheetAlreadyInCircle = ixToRemove > 0
    If bSheetAlreadyInCircle Then
        RemoveActiveSheetFromCircle
    Else
        AddActiveSheetToCircle
    End If
End Sub


' see work.xlsb, uses Form

'Sub FlashTargetSheetName(sTarget As String)
'    Dim frm As New TargetSheetNameForm
'    frm.Label1 = sTarget
'    frm.Show vbModeless
'    DoEvents
'    Sleep 600
'    Unload frm
'    Set frm = Nothing
'End Sub


Sub JumpToNextSheetInCircle()
    Application.StatusBar = False
    
    Dim rCircle As Range: Set rCircle = CircleSheetsRange()
    
    Dim nSheetsInCircle As Integer: nSheetsInCircle = Application.WorksheetFunction.CountA(rCircle)
    If nSheetsInCircle = 0 Then
    
        MsgBox "no sheets in circle - add w/ Ctrl-Q"
    
    Else
            
        Dim sTarget As String
        
        Dim ixCurrent As Integer: ixCurrent = match(ActiveSheet.Name, rCircle)
        If ixCurrent = 0 Then
        
        
            ' we are not on a circle sheet - go to first one
            sTarget = rCircle.Cells(1)
        
        Else
        
            ' we are in the circle
            
            If ixCurrent = nSheetsInCircle Then
            
                ' on the last circle member - go to first
                sTarget = rCircle.Cells(1)
            
            Else
            
                sTarget = rCircle.Cells(ixCurrent + 1)
            End If
            
        End If
        
        If SheetExists(sTarget) Then
            'FlashTargetSheetName sTarget
            Worksheets(sTarget).Activate
        Else
            RemoveSheetFromCircle sTarget, sTarget & " not found - removed"
        End If
    End If
End Sub


Sub JumpToPrevSheetInCircle()
    Dim rCircle As Range: Set rCircle = CircleSheetsRange()
    
    Dim nSheetsInCircle As Integer: nSheetsInCircle = Application.WorksheetFunction.CountA(rCircle)
    If nSheetsInCircle = 0 Then
    
        MsgBox "no sheets in circle - add w/ Ctrl-Q"
    
    Else
            
        Dim sTarget As String
        
        Dim ixCurrent As Integer: ixCurrent = match(ActiveSheet.Name, rCircle)
        If ixCurrent = 0 Then
        
        
            ' we are not on a circle sheet - go to first one
            sTarget = rCircle.Cells(nSheetsInCircle)
        
        Else
        
            ' we are in the circle
            
            If ixCurrent = 1 Then
            
                ' on the last circle member - go to first
                sTarget = rCircle.Cells(nSheetsInCircle)
                        
            Else
            
                sTarget = rCircle.Cells(ixCurrent - 1)
            End If
            
        End If
        
        Worksheets(sTarget).Activate
    End If
End Sub



'>> Sheets

' ((EPMHNMQ)) adjust EXCEL_LAST_ROW_PJ for pre-10 versions

'*************************************
' importing
'*************************************

Function ImportSheet(sWorkbook As String, sExportSheet As String, sImportSheet As String) As Boolean
    SetSilentApplicationState
    On Error GoTo err
    
    Dim wsImport As Worksheet: Set wsImport = SafeGetWorksheet(sImportSheet)
    If wsImport Is Nothing Then MsgError "no import sheet '" & sImportSheet & "'"
    
    Dim wbExport As Workbook: Set wbExport = SafeGetWorkbook(sWorkbook)
    If wbExport Is Nothing Then MsgError "open '" & sWorkbook & "' first"
    
    Dim wsExport As Worksheet: Set wsExport = SafeGetWorksheet(sExportSheet, wbExport)
    If wsExport Is Nothing Then MsgError "no export sheet '" & sExportSheet & "'"
    
    ShowStatus SubstituteParams("importing {1} from {2} ...", sImportSheet, sExportSheet)
    
    ' first unmerge anything on the target sheet, else we cannot paste (even if source has the same merged fields)
    wsImport.Cells.UnMerge
        
    ' copy + paste
    wsExport.Cells.Copy
    wsImport.Cells.PasteSpecial xlPasteValues
    wsImport.Cells.PasteSpecial xlPasteFormats
    
    ColorAsImportSheet wsImport
    
    ImportSheet = True
    GoTo done
    
err:
    ImportSheet = False
    
done:
    RevertApplicationState
End Function


'*************************************
' getters
'*************************************

Function SafeGetWorkbook(sWorkbook As String) As Workbook
    On Error Resume Next
    Set SafeGetWorkbook = Workbooks(sWorkbook)
End Function


Function SafeGetWorksheet(sWorksheet, Optional wb As Workbook) As Worksheet
    On Error Resume Next
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Set SafeGetWorksheet = wb.Worksheets(sWorksheet)
End Function


Public Function GetWorksheetName(r As Range) As String
    GetWorksheetName = r.Worksheet.Name
End Function


Public Function GetWorksheetNameVolatile(r As Range) As String
    Application.Volatile
    GetWorksheetNameVolatile = GetWorksheetName(r)
End Function


Function SheetExists(sName, Optional wb As Workbook) As Boolean
    On Error Resume Next
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim wsExisting As Worksheet: Set wsExisting = wb.Worksheets(sName)
    SheetExists = Not (wsExisting Is Nothing)
End Function

Function WorksheetExists(sName, Optional wb As Workbook) As Boolean
    WorksheetExists = SheetExists(sName, wb)
End Function


Function CountWorksheets()
    CountWorksheets = Sheets.Count
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


'*************************************
' OrderSheetNames
'*************************************

Sub OrderSheetNames()
    ' goal: order sheets (below Microsoft Excel Objekte) so that their order is the same as the visible one
    ' call from VBA, but outside the project which sheets we want to order

    Dim varItem As Variant

    Dim index As New Dictionary
    Dim ixWorksheet As Integer
    Dim ws As Worksheet
    For Each ws In Worksheets
        ixWorksheet = ixWorksheet + 1
        index.Add ws.Name, ixWorksheet
    Next

    Dim bIstTabelle As Boolean
    Dim iNumber As Integer
    
    For Each varItem In ActiveWorkbook.VBProject.VBComponents
        'Type 100 is a worksheet
        Dim bIsWorksheet As Boolean: bIsWorksheet = varItem.Type = 100
        Dim bIsNormalSheet As Boolean: bIsNormalSheet = varItem.Name <> "ThisWorkbook" And varItem.Name <> "DieseArbeitsmappe" And varItem.Name <> "Tabelle00"
        If bIsWorksheet And bIsNormalSheet Then
            Dim sCodeName As String: sCodeName = varItem.Name
            bIstTabelle = Left(sCodeName, 7) = "Tabelle"
            If bIstTabelle Then
                'iNumber = Val(Mid(sCodename, 8))
                iNumber = iNumber + 1
                varItem.Name = "Tabelle" & Format(iNumber, "0000")
            End If
        End If
    Next

    ' 24.02.18 problems w/ varItem.Properties("Name") for multiple sheets looking like "DieseArbeitsmappe"
    ' also Automatisierungsfehler in Is64Bit()
    On Error Resume Next

    For Each varItem In ActiveWorkbook.VBProject.VBComponents
        'Type 100 is a worksheet
        bIsWorksheet = varItem.Type = 100
        bIsNormalSheet = varItem.Name <> "ThisWorkbook" And varItem.Name <> "DieseArbeitsmappe" And varItem.Name <> "Tabelle00"
        If bIsWorksheet And bIsNormalSheet Then
            Dim sName As String: sName = varItem.Properties("Name").value
            If sName <> "" Then
                sCodeName = varItem.Name
                bIstTabelle = Left(sCodeName, 7) = "Tabelle"
                If bIstTabelle Then
                    ixWorksheet = index(sName)
                    iNumber = ixWorksheet
                    varItem.Name = "Tabelle" & Format(iNumber, "000")
                End If
            End If
        End If
    Next
End Sub



'>> Sort

' ((UNKLAKD)) merge Sorted() and SortVector(), see also Unique...


'*************************************
' collections + dictionaries
'*************************************

' ((UNKLAKD)) merge Sorted() and SortVector(), see also Unique...
Function Sorted(c As Collection, Optional bStrings As Boolean = False) As Collection
    If c.Count = 0 Then
        Set Sorted = c
    Else
        Dim av()
        av = CollectionToArray(c)
        If bStrings Then
            QuickSortStrings av, LBound(av), UBound(av)
        Else
            QuickSort av, LBound(av), UBound(av)
        End If
        Dim ix
        Set Sorted = New Collection
        For ix = 1 To UBound(av)
            Sorted.Add av(ix)
        Next ix
    End If
End Function

Function SortedCollection(cUnsorted As Collection, Optional bStrings As Boolean = False) As Collection
    Set SortedCollection = Sorted(cUnsorted, bStrings)
End Function

Function SortedDictionary(dictUnsorted As Dictionary, Optional bStrings As Boolean = False) As Dictionary
    Dim cUnsortedKeys As New Collection
    Dim v
    For Each v In dictUnsorted
        cUnsortedKeys.Add v
    Next
    
    Dim cSortedKeys As Collection: Set cSortedKeys = Sorted(cUnsortedKeys, bStrings)
    
    Set SortedDictionary = New Dictionary
    For Each v In cSortedKeys
        SortedDictionary.Add v, dictUnsorted(v)
    Next
End Function


Function SafeGetArray(v) As Variant()
    Dim av()
    If TypeOf v Is Range Then
        Dim r As Range: Set r = v
        av = RangeToVector(r)
    Else
        av = v
        Debug.Assert IsVector(av)
        MakeVector av
    End If
    SafeGetArray = av
End Function


'*************************************
' helpers for vector/array management
'*************************************

Function SortVector(v, Optional sEmpty As String) As Variant()
    Dim av(): av = SafeGetArray(v)
        
    If sEmpty <> "" Then
        Dim ix As Long
        For ix = LBound(av) To UBound(av)
            If IsEmpty(av(ix)) Then av(ix) = sEmpty
        Next
    End If
    
    QuickSort av, LBound(av), UBound(av)
    
    SortVector = AutoOrientVector(av)
End Function


Function ReverseVector(v) As Variant()
    Dim av(): av = SafeGetArray(v)

    av = ReverseArray(av)
    ReverseVector = AutoOrientVector(av)
End Function


Function ReverseArrayWithArrayList(arr As Variant) As Variant
    ' https://stackoverflow.com/questions/40563940/vba-reverse-an-array
    ' Problem: works like a charm - - but always 0-based return value
    
    Dim val As Variant

    With CreateObject("System.Collections.ArrayList") '<-- create a "temporary" array list with late binding
        For Each val In arr '<--| fill arraylist
            .Add val
        Next val
        .Reverse '<--| reverse it
        ReverseArrayWithArrayList = .Toarray '<--| write it into an array
    End With
End Function


Function ReverseArray(InputArray As Variant) As Variant
    ' https://stackoverflow.com/questions/40563940/vba-reverse-an-array
    ' better than above, because we can also work w/ 1-based arrays
    
    Dim Ndx2 As Long: Ndx2 = UBound(InputArray)
    ' loop from the LBound of InputArray to the midpoint of InputArray
    Dim Ndx As Long
    For Ndx = LBound(InputArray) To ((UBound(InputArray) - LBound(InputArray) + 1) \ 2)
        'swap the elements
        Dim Temp: Temp = InputArray(Ndx)
        InputArray(Ndx) = InputArray(Ndx2)
        InputArray(Ndx2) = Temp
        ' decrement the upper index
        Ndx2 = Ndx2 - 1
    Next Ndx
    
    ReverseArray = InputArray
End Function


Function TopX(v, nTop As Integer) As Variant()
    If nTop <= 0 Then
        TopX = v
    Else
        Dim av(): av = SafeGetArray(v)
        Dim avTop(): ReDim avTop(1 To nTop)
        Dim ixTop As Integer
        For ixTop = 1 To nTop
            avTop(ixTop) = av(ixTop)
        Next
    End If
    TopX = AutoOrientVector(avTop)
End Function


Function CreateIndexArray(n As Long, Optional iLBound As Long = 1) As Variant
    Dim iUBound As Long: iUBound = n - (iLBound - 1)
    Dim a(): ReDim a(iLBound To iUBound)
    Dim ix As Long
    For ix = iLBound To iUBound
        a(ix) = ix
    Next
    CreateIndexArray = a
End Function


'*************************************
' QuickSort
'*************************************

Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
    Dim Pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long

    tmpLow = inLow
    tmpHi = inHi

    Pivot = vArray((inLow + inHi) \ 2)

    While (tmpLow <= tmpHi)

        While (vArray(tmpLow) < Pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend

        While (Pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend

        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend

    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub


Public Sub QuickSortStrings(vArray As Variant, inLow As Long, inHi As Long)
    Dim Pivot   As String
    Dim tmpSwap As String
    Dim tmpLow  As Long
    Dim tmpHi   As Long

    tmpLow = inLow
    tmpHi = inHi

    Pivot = vArray((inLow + inHi) \ 2)

    While (tmpLow <= tmpHi)

        While (StrComp(vArray(tmpLow), Pivot, vbTextCompare) = -1) And (tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend

        While (StrComp(Pivot, vArray(tmpHi), vbTextCompare) = -1) And (tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend

        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend

    If (inLow < tmpHi) Then QuickSortStrings vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSortStrings vArray, tmpLow, inHi
End Sub


Public Sub QuickSortWithIndex(vArray As Variant, vIndexArray As Variant, inLow As Long, inHi As Long)

    ' sorts vArray (as normal QuickSort() above) but also swaps vIndexArray in parallel

    Dim Pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long

    tmpLow = inLow
    tmpHi = inHi

    Pivot = vArray((inLow + inHi) \ 2)

    While (tmpLow <= tmpHi)

        While (vArray(tmpLow) < Pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        While (Pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend

        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            
            tmpSwap = vIndexArray(tmpLow)
            vIndexArray(tmpLow) = vIndexArray(tmpHi)
            vIndexArray(tmpHi) = tmpSwap
            
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend

    If (inLow < tmpHi) Then QuickSortWithIndex vArray, vIndexArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSortWithIndex vArray, vIndexArray, tmpLow, inHi

End Sub



'>> Speech

Private Sub ChangeVoiceDemo()
    TextToSpeech "Excel spricht mit mir.", VOICE_DE, 2, 50
    TextToSpeech "Excel is talking to me.", VOICE_EN_F, 2, 5
    TextToSpeech "Excel spricht mit mir.", VOICE_DE, -10, 30
    TextToSpeech "Excel is talking to me.", VOICE_EN_M, 10, 100
End Sub


Public Sub TextToSpeech(Words As String, Optional Person As Long = VOICE_EN_M, Optional Rate As Long = 0, Optional Volume As Long = 100)
    Dim voc As New SpeechLib.SpVoice
    
    With voc
        Set .Voice = .GetVoices.Item(Person)
        .Rate = Rate
        .Volume = Volume
        .Speak Words
    End With
End Sub


Private Sub DumpAvailableVoices()
    Dim voc As New SpeechLib.SpVoice
    Debug.Print voc.GetVoices.Count & " available voices:"
    Dim i As Long
    For i = 0 To voc.GetVoices.Count - 1
        Set voc.Voice = voc.GetVoices.Item(i)
        Debug.Print " " & i & " - " & voc.Voice.GetDescription
        voc.Speak "test audio"
    Next i
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

Sub ResetEverything()
    ClearStatusBar
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    WorksheetChangeDisabled = False
    ForceInitApplicationStates
End Sub



'>> Strand

' 31.10.24 added from QA

'*************************************
' strand helpers
'*************************************

Function CreateStrand(Optional bPutOnClipboard As Boolean = True)
    Randomize
    Dim sStrand As String: sStrand = LCase(RandomString(STRAND_SIZE))
    PutOnClipboard sStrand
    CreateStrand = sStrand
End Function
    

Function GetStrandFromString(s) As String
    GetStrandFromString = RegMatch(s, PAT_Strand, 0)
End Function


Function GetStrandFromActiveComment() As String
    Debug.Assert OnAnsarada()
    Dim rCurrentRow As Range: Set rCurrentRow = Selection.EntireRow
    Dim sComment As String: sComment = rCurrentRow.Cells(COL_INDEX_Comment)
    GetStrandFromActiveComment = GetStrandFromString(sComment)
End Function


Sub WrapStrandOnClipboardInGrayFont()
    ' 12.12.21 not used
    Dim sStrand As String: sStrand = RegMatch(GetFromClipboard(), "^([A-Za-z][A-Za-z][A-Za-z][A-Za-z][A-Za-z][A-Za-z][A-Za-z]?)$", 1)
    If sStrand <> "" Then
        Dim sFormattedStrand As String: sFormattedStrand = FS("<font color=\#F0F0F0>[#]</font>", sStrand)
        PutOnClipboard sFormattedStrand
    End If
End Sub


Function RemoveStrandFromHtml(sHtml As String) As String
    sHtml = RegReplace(sHtml, PAT_Strand & " *<br/>", "")
    sHtml = RegReplace(sHtml, PAT_Strand, "")
    RemoveStrandFromHtml = sHtml
End Function


'*************************************
' extract multiple strands from string
'*************************************

Sub TestStrandsFromString()
    Dim s: s = "asdf [aueaue] wert [wertew] etrz"
    Dim c As Collection: Set c = GetStrandsFromString(s)
    Debug.Assert c(1) = "[aueaue]"
    Debug.Assert c(2) = "[wertew]"
End Sub


Function GetStrandsFromString(s) As Collection
    Dim cStrands As New Collection
    locGetStrandsFromString s, cStrands
    Set GetStrandsFromString = cStrands
End Function


Sub locGetStrandsFromString(s, cStrands As Collection)
    Dim sStrand: sStrand = GetStrandFromString(s)
    If sStrand = "" Then
        Exit Sub
    Else
        cStrands.Add sStrand
        Dim ixFound: ixFound = InStr(1, s, sStrand)
        s = Mid(s, ixFound + Len(sStrand))
        locGetStrandsFromString s, cStrands
    End If
End Sub



'>> Strings

' 02.07.15 baselined
' 13.07.15 added EndsWith
' 15.07.15 case insensitive matching
' 15.07.15 fixed ucase, TextCompare

Function Quoted(s) As String
    Quoted = """" & s & """"
End Function


Function PadRight(s As String, n As Integer, sPadChar As String) As String
    If Len(s) = n Then
        PadRight = s
    Else
        If n < Len(s) Then
            PadRight = Left(s, n)
        Else
            PadRight = s & RepeatString(sPadChar, n - Len(s))
        End If
    End If
End Function


Function PadLeft(s As String, n As Integer, sPadChar As String) As String
    If Len(s) = n Then
        PadLeft = s
    Else
        If n < Len(s) Then
            PadLeft = Right(s, n)
        Else
            PadLeft = RepeatString(sPadChar, n - Len(s)) & s
        End If
    End If
End Function


Function TPct(d As Double, Optional nDigits As Integer) As String
    Dim sFormat As String: sFormat = "0%"
    If nDigits > 0 Then sFormat = sFormat & "." & RepeatString("0", nDigits)
    TPct = Format(d, sFormat)
End Function


Function TrimAll(s As String) As String
    TrimAll = Trim(Replace(s, Chr(10), " "))
End Function


Function AsString(v As Variant) As String
    If IsError(v) Then
        AsString = "#NV"
    Else
        On Error Resume Next
        AsString = v
    End If
End Function


Function MatchesOneOf(sSearch As String, r As Range) As Boolean
    MatchesOneOf = match(sSearch, r) > 0
End Function


'*************************************
' string -> array
'*************************************

Function StringToArray(s As Variant) As Variant()
    Dim a(): ReDim a(1 To Len(s))
    Dim ix
    For ix = 1 To Len(s)
        a(ix) = Mid(s, ix, 1)
    Next
    StringToArray = a
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
' formatting
'*************************************

Public Function RandomString(Optional n As Integer) As String
    If n <= 0 Then n = 7
    Dim ix As Integer
    Dim s As String
    For ix = 1 To n
        ' 65 + 26 would be [, 65 + 25 = Z
        s = s + Chr(65 + Rnd() * 25)
    Next ix
    RandomString = s
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


Function Contains(sTarget As String, sCheck As String, Optional bCaseSensitive As Boolean) As Boolean
    If bCaseSensitive Then
        Contains = InStr(1, sTarget, sCheck) <> 0
    Else
        Contains = InStr(1, UCase(sTarget), UCase(sCheck)) <> 0
    End If
End Function


Function TextCompare(s1, s2) As Integer
    TextCompare = StrComp(s1, s2, vbTextCompare)
End Function

Function StringsEqual(s1 As String, s2 As String) As Boolean
    StringsEqual = TextCompare(s1, s2) = 0
End Function


'*************************************
' concat
'*************************************

Function ConcatenateRangeWithSep(r As Range, Optional vSep As Variant) As String
    Dim sSep As String
    If IsMissing(vSep) Then sSep = "" Else sSep = vSep
    Dim rCell As Range
    For Each rCell In r
        If ConcatenateRangeWithSep <> "" Then ConcatenateRangeWithSep = ConcatenateRangeWithSep & sSep
        ConcatenateRangeWithSep = ConcatenateRangeWithSep & rCell
    Next
End Function

Function ConcatenateWithSep(v As Variant, Optional vSep As Variant) As String
    Dim r As Range
    If TypeOf v Is Range Then
        Set r = v
        ConcatenateWithSep = ConcatenateRangeWithSep(r, vSep)
    Else
        Dim sSep As String
        If IsMissing(vSep) Then sSep = "" Else sSep = vSep
        If IsArray(v) Then
        
            Dim ix As Integer
            Dim a(): a = v
            For ix = LBound(a) To UBound(a)
                If TypeOf a(ix) Is Range Then
                    Set r = a(ix)
                    ConcatenateWithSep = ConcatenateRangeWithSep(r, vSep)
                Else
                    If ConcatenateWithSep <> "" Then ConcatenateWithSep = ConcatenateWithSep & sSep
                    Dim s As String: s = a(ix)
                    ConcatenateWithSep = ConcatenateWithSep & s
                End If
            Next ix
        
        Else
            ConcatenateWithSep = v
        End If
    End If
End Function


Function ConcatenateRangeWithSep2(r As Range, Optional vSep As Variant) As String
    Dim sSep As String
    If IsMissing(vSep) Then sSep = "" Else sSep = vSep
    Dim rCell As Range
    For Each rCell In r
        If ConcatenateRangeWithSep2 <> "" Then ConcatenateRangeWithSep2 = ConcatenateRangeWithSep2 & sSep
        ConcatenateRangeWithSep2 = ConcatenateRangeWithSep2 & IfError(rCell, "err")
    Next
End Function

Function ConcatenateRangeWithSep3(r As Range, Optional vSep As Variant) As String
    Dim sSep As String
    If IsMissing(vSep) Then sSep = "" Else sSep = vSep
    Dim rCell As Range
    For Each rCell In r
        If ConcatenateRangeWithSep3 <> "" Then ConcatenateRangeWithSep3 = ConcatenateRangeWithSep3 & sSep
        ConcatenateRangeWithSep3 = ConcatenateRangeWithSep3 & SafeConcatenateValue(rCell)
    Next
End Function

Function SafeConcatenateValue(rCell As Range) As String
    If IsError(rCell) Then
        If rCell = ErrNA() Then
            SafeConcatenateValue = "#NA"
        ElseIf rCell = ErrValue() Then
            SafeConcatenateValue = "#VALUE"
        Else
            SafeConcatenateValue = "#?"
        End If
    Else
        If IsNumeric(rCell) Then
            SafeConcatenateValue = AsString(rCell)
        Else
            Dim s As String: s = rCell.Value2
            If Left(s, 1) = "#" Then
                SafeConcatenateValue = "#error"
            Else
                SafeConcatenateValue = s
            End If
        End If
    End If
End Function


'*************************************
' split
'*************************************

Function SplitText(s As String, Optional sSep As String) As Variant
    If sSep = "" Then sSep = " "
    Dim a: a = Split(s, sSep)
    Dim v(): ReDim v(UBound(a))
    Dim ix As Integer
    For ix = LBound(a) To UBound(a)
        If IsNumeric(a(ix)) Then
            v(ix) = a(ix) + 0
        Else
            v(ix) = a(ix)
        End If
    Next ix
    SplitText = v
End Function


Function SplitMultiDelims(Text As String, DelimChars As String) As String()
    ' see http://www.cpearson.com/excel/splitondelimiters.aspx
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SplitMutliChar
    ' This function splits Text into an array of substrings, each substring
    ' delimited by any character in DelimChars. Only a single character
    ' may be a delimiter between two substrings, but DelimChars may
    ' contain any number of delimiter characters. If you need multiple
    ' character delimiters, use the SplitMultiDelimsEX function. It returns
    ' an unallocated array it Text is empty, a single element array
    ' containing all of text if DelimChars is empty, or a 1 or greater
    ' element array if the Text is successfully split into substrings.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Pos1 As Long
    Dim n As Long
    Dim m As Long
    Dim arr() As String
    Dim i As Long
    
    ''''''''''''''''''''''''''''''''
    ' if Text is empty, get out
    ''''''''''''''''''''''''''''''''
    If Len(Text) = 0 Then
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' if DelimChars is empty, return original text
    '''''''''''''''''''''''''''''''''''''''''''''
    If DelimChars = vbNullString Then
        SplitMultiDelims = Array(Text)
        Exit Function
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' oversize the array, we'll shrink it later so
    ' we don't need to use Redim Preserve
    '''''''''''''''''''''''''''''''''''''''''''''''
    ReDim arr(1 To Len(Text))
    
    i = 0
    n = 0
    Pos1 = 1
    
    For n = 1 To Len(Text)
        For m = 1 To Len(DelimChars)
            If StrComp(Mid(Text, n, 1), Mid(DelimChars, m, 1), vbTextCompare) = 0 Then
                i = i + 1
                arr(i) = Mid(Text, Pos1, n - Pos1)
                Pos1 = n + 1
                n = n + 1
            End If
        Next m
    Next n
    
    If Pos1 <= Len(Text) Then
        i = i + 1
        arr(i) = Mid(Text, Pos1)
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' chop off unused array elements
    ''''''''''''''''''''''''''''''''''''''
    ReDim Preserve arr(1 To i)
    SplitMultiDelims = arr
        
End Function


'>> System

Public Function Is64Bit() As Boolean
#If VBA7 And Win64 Then
    Is64Bit = True
#Else
    Is64Bit = False
#End If
End Function

Sub YamYam()
    msgbo GetMemUsage()
End Sub

Function GetMemUsage()
    Dim objSWbemServices: Set objSWbemServices = GetObject("winmgmts:")
    GetMemUsage = objSWbemServices.Get("Win32_Process.Handle='" & GetCurrentProcessId & "'").WorkingSetSize / 1024
    Set objSWbemServices = Nothing
End Function



'>> Tapestry

Function ParseRange(sRange As String) As Range
    sRange = Replace(sRange, "'", "")
    
    Const RE1 = "(^\$?[A-Z]+\$?[0-9]+$)"
    Dim sAddress As String: sAddress = RegMatch(sRange, RE1, 1)
    If sAddress <> "" Then
        Set ParseRange = ActiveSheet.Range(sAddress)
    Else
    
        Const RE2 = "(\[([^]]+)\])?([^!]+)?!?(.+)"
        Dim sWorkbook As String: sWorkbook = RegMatch(sRange, RE2, 2)
        Dim sWorksheet As String: sWorksheet = RegMatch(sRange, RE2, 3)
        sAddress = RegMatch(sRange, RE2, 4)
        
        Dim wb As Workbook
        If sWorkbook = "" Then
            Set wb = ActiveWorkbook
        Else
            Set wb = Workbooks(sWorkbook)
        End If
        
        Dim ws As Worksheet
        If sWorksheet = "" Then
            Set ws = wb.ActiveSheet
        Else
            Set ws = wb.Worksheets(sWorksheet)
        End If
        
        Set ParseRange = ws.Range(sAddress)
    End If
End Function


Sub ConvertIndexToDirectLink()
    SetSilentApplicationState
    On Error GoTo done
    
    Dim sWrapperFunction As String: sWrapperFunction = InputBox("Wrapper Function (e.g. ExtractPdfNumber)")
    If sWrapperFunction = "" Then
        If MsgBox("Really no wrapper, or cancel?", vbYesNoCancel) <> vbYes Then
            Exit Sub
        End If
    End If

    Dim rToConvert As Range: Set rToConvert = Selection
    
    Dim rCellToConvert As Range
    For Each rCellToConvert In rToConvert
    
        If rCellToConvert.HasFormula Then
    
            '=ExtractPdfNumber(INDEX('2013 P&L Details'!$B:$B,MATCH($B14,'2013 P&L Details'!$A:$A,0)))
            'Dim vTarget: vTarget = "=ExtractPdfNumber(INDEX('2013 P&L Details'!$B:$B,MATCH($B14,'2013 P&L Details'!$A:$A,0)))"
            Dim vTarget: vTarget = rCellToConvert.formula
            
            Const RE1 = "INDEX\(([^,]+), ?([^)]+)\)"
            Dim sIndex As String: sIndex = RegMatch(vTarget, RE1, 1)
            Dim sMatch As String: sMatch = RegMatch(vTarget, RE1, 2)
            
            If sIndex = "" Then
            
                ' skip, no formula for us
            
            Else
            
                Dim rIndex As Range: Set rIndex = ParseRange(sIndex)
                Dim rMatch As Range: Set rMatch = ParseRange(sMatch)
                If IsError(rMatch) Then
                    rCellToConvert = ErrNA()
                Else
                    Dim ixMatch As Long: ixMatch = rMatch
                    Dim rFound As Range: Set rFound = rIndex(ixMatch)
                    
                    Dim sNewFormula As String: sNewFormula = FS("=#(#)", sWrapperFunction, GetAddress(rFound, True))
                    'Debug.Print sNewFormula
                    
                    rCellToConvert.formula = sNewFormula
                End If
            End If
        End If
    Next

done:
    RevertApplicationState
End Sub



Sub ConvertIndexMatchToDirectLink()
    SetSilentApplicationState
    On Error GoTo done
    
    Dim sWrapperFunction As String: sWrapperFunction = InputBox("Wrapper Function (e.g. ExtractPdfNumber)")
    If sWrapperFunction = "" Then
        If MsgBox("Really no wrapper, or cancel?", vbYesNoCancel) <> vbYes Then
            Exit Sub
        End If
    End If
    
    Dim bRetainFormula As Boolean: bRetainFormula = MsgBox("Retain formula if no match?", vbYesNo) = vbYes

    Dim rToConvert As Range: Set rToConvert = Selection
    
    Dim rCellToConvert As Range
    For Each rCellToConvert In rToConvert
    
        If rCellToConvert.HasFormula Then
    
            '=ExtractPdfNumber(INDEX('2013 P&L Details'!$B:$B,MATCH($B14,'2013 P&L Details'!$A:$A,0)))
            'Dim vTarget: vTarget = "=ExtractPdfNumber(INDEX('2013 P&L Details'!$B:$B,MATCH($B14,'2013 P&L Details'!$A:$A,0)))"
            Dim sFormula: sFormula = rCellToConvert.formula
            
            Const RE1 = "(INDEX\(([^,]+), ?MATCH\(([^,]+),([^,]+),0\)\))"
            Dim sMatchedIndexMatch As String: sMatchedIndexMatch = RegMatch(sFormula, RE1, 1)
            
            If sMatchedIndexMatch = "" Then
            
                ' skip, no formula for us
            
            Else
            
                Dim sIndex As String: sIndex = RegMatch(sFormula, RE1, 2)
                
                Dim sFind As String: sFind = RegMatch(sFormula, RE1, 3)
                Dim sMatch As String: sMatch = RegMatch(sFormula, RE1, 4)
                
                If sFind = "" Or sMatch = "" Then
                
                    ' somehow a partial match - - nothing for us
                    
                Else
                
                    Const RE2 = "([^!]+)!(.+)"
                    
                    ' TODO: sMatchSheet und sIndexSheet können entfernt werden, s. oben
                    
                    'Dim sIndexSheet As String: sIndexSheet = Replace(RegMatch(sIndex, RE2, 1), "'", "")
                    'Dim sIndexRange As String: sIndexRange = RegMatch(sIndex, RE2, 2)
                    'Dim sMatchSheet As String: sMatchSheet = Replace(RegMatch(sMatch, RE2, 1), "'", "")
                    'Dim sMatchRange As String: sMatchRange = RegMatch(sMatch, RE2, 2)
                    
                    Dim vFind: vFind = Range(sFind)
                    
                    'Set rIndex = Worksheets(sIndexSheet).Range(sIndexRange)
                    'Dim rMatch As Range: Set rMatch = Worksheets(sMatchSheet).Range(sMatchRange)
                                        
                    Dim rIndex As Range: Set rIndex = ParseRange(sIndex)
                    Dim rMatch As Range: Set rMatch = ParseRange(sMatch)
                    
                    
                    Dim ixMatch As Integer: ixMatch = match(vFind, rMatch)
                    
                    If ixMatch = 0 Then
                    
                        If Not bRetainFormula Then rCellToConvert.ClearContents
                    
                    Else
                    
                        Dim rFound As Range: Set rFound = rIndex(ixMatch)
                        Dim sNewFormula As String
                        
                        If sWrapperFunction = "" Then
                            sNewFormula = Replace(sFormula, sMatchedIndexMatch, GetAddress(rFound, True))
            
                        Else
                            sNewFormula = FS("=#(#)", sWrapperFunction, GetAddress(rFound, True))
                            'Debug.Print sNewFormula
                        End If
                        
                        rCellToConvert.formula = sNewFormula
                    
                    End If
                End If
            End If
        End If
    Next

done:
    RevertApplicationState
End Sub



Sub ConvertMatchToDirectLink()
    SetSilentApplicationState
    On Error GoTo done
    
    Dim sWrapperFunction As String: sWrapperFunction = InputBox("Wrapper Function")
    If sWrapperFunction = "" Then
        If MsgBox("Really no wrapper, or cancel?", vbYesNoCancel) <> vbYes Then
            Exit Sub
        End If
    End If
    
    Dim bRetainFormula As Boolean: bRetainFormula = MsgBox("Retain formula if no match?", vbYesNo) = vbYes

    Dim rToConvert As Range: Set rToConvert = Selection
    
    Dim rCellToConvert As Range
    For Each rCellToConvert In rToConvert
    
        If rCellToConvert.HasFormula Then
    
            Dim sFormula: sFormula = rCellToConvert.formula
            
            Const RE1 = "(MATCH\(([^,]+),([^,]+),0\))"
            Dim sMatchedIndexMatch As String: sMatchedIndexMatch = RegMatch(sFormula, RE1, 1)
            
            If sMatchedIndexMatch = "" Then
            
                ' skip, no formula for us
            
            Else
            
                Dim sFind As String: sFind = RegMatch(sFormula, RE1, 2)
                Dim sMatch As String: sMatch = RegMatch(sFormula, RE1, 3)
                
                If sFind = "" Or sMatch = "" Then
                
                    ' somehow a partial match - - nothing for us
                    
                Else
            
                    Dim vFind: vFind = Evaluate(sFind)
                    Dim rMatch As Range: Set rMatch = ParseRange(sMatch)
                    
                    Dim ixMatch As Integer: ixMatch = match(vFind, rMatch)
                    
                    If ixMatch = 0 Then
                    
                        If Not bRetainFormula Then rCellToConvert.ClearContents
                    
                    Else
                    
                        Dim rFound As Range: Set rFound = rMatch(ixMatch)
                        Dim sNewFormula As String
                        
                        If sWrapperFunction = "" Then
                            sNewFormula = Replace(sFormula, sMatchedIndexMatch, GetAddress(rFound, True))
            
                        Else
                            sNewFormula = FS("=#(#)", sWrapperFunction, GetAddress(rFound, True))
                            'Debug.Print sNewFormula
                        End If
                        
                        rCellToConvert.formula = sNewFormula
                    
                    End If
                End If
            End If
        End If
    Next

done:
    RevertApplicationState
End Sub



Function SlidingMatch(v, r As Range, Optional vStart)
    Dim ixStart As Integer
    
    If TypeOf vStart Is Range Then
        Dim rStarts As Range: Set rStarts = vStart
        
        ' search through "past" matches (upwards)
        Dim ix As Integer
        For ix = rStarts.Cells.Count To 1 Step -1
            Dim rStart As Range: Set rStart = rStarts(ix)
            If rStart.Value2 = 0 Then
                ' continue looking for the first match, upwards
            Else
                ' found a non-0
                ixStart = rStart.Value2
                Exit For
            End If
        Next
    Else
        ixStart = vStart
    End If
    
    If ixStart = 0 Then ixStart = 1

    Debug.Assert IsVectorRange(r)
    
    Dim rSliding As Range: Set rSliding = r.Parent.Range(r.Cells(ixStart + 1), r.Cells(r.Cells.Count))
    Dim ixMatch As Integer: ixMatch = match(v, rSliding)
    If ixMatch = 0 Then
        ' indicate that we haven't found any match
        SlidingMatch = 0
    Else
        SlidingMatch = ixMatch + ixStart
    End If
End Function



'>> Timer

Sub StartTimer()
    g_dTimer = MicroTimer
End Sub

Function TimerElapsed() As Double
    TimerElapsed = MicroTimer - g_dTimer
End Function

Function TimerElapsedStr(Optional sFormat As String = "0.000") As String
    Dim dElapsed As Double: dElapsed = TimerElapsed
    TimerElapsedStr = Format(dElapsed, sFormat)
End Function


Sub TestTimer()
    Dim d As Double: d = MicroTimer()
    Sleep 100
    MsgBox MicroTimer - d
    
    MsgBox Is64Bit
End Sub

Public Function MicroTimer() As Double
    ' https://fastexcel.wordpress.com/2011/10/26/match-vs-find-vs-variant-array-vba-performance-shootout/
    ' returns seconds
    ' uses Windows API calls to the high resolution timer
    '
    Dim cyTicks1 As Currency
    Dim cyTicks2 As Currency
    Static cyFrequency As Currency
    '
    MicroTimer = 0
    '
    ' get frequency
    '
    If cyFrequency = 0 Then getFrequency cyFrequency
    '
    ' get ticks
    '
    getTickCount cyTicks1
    getTickCount cyTicks2
    If cyTicks2 < cyTicks1 Then cyTicks2 = cyTicks1
    '
    ' calc seconds
    '
    If cyFrequency Then MicroTimer = cyTicks2 / cyFrequency
End Function



'>> Timeseries

' 02.07.15 baselined

Function TTM(r As Range) As Double
    Dim rTTM As Range: Set rTTM = r.Worksheet.Range(r, r.Offset(0, -11))
    TTM = Application.WorksheetFunction.sum(rTTM)
End Function

Function TnM(r As Range, nMonths As Integer, Optional rFirst As Range, Optional rSync As Range) As Double
    Dim nMonthCols As Integer: nMonthCols = nMonths
    Dim mult As Double: mult = 1
    If Not rFirst Is Nothing Then
        Dim nMaxMonthCols As Integer
        If Not rSync Is Nothing Then
            nMaxMonthCols = rSync.Column - rFirst.Column + 1
        Else
            nMaxMonthCols = r.Column - rFirst.Column + 1
        End If
        If nMonthCols > nMaxMonthCols Then
            mult = nMonthCols / nMaxMonthCols
            nMonthCols = nMaxMonthCols
        End If
    End If
    Dim offs As Integer: offs = -(nMonthCols - 1)
    Dim rTnM As Range: Set rTnM = r.Worksheet.Range(r, r.Offset(0, offs))
    Dim sum As Double: sum = Application.WorksheetFunction.sum(rTnM)
    TnM = sum * mult
End Function


Function CentralSum(r As Range, nMonths As Integer, Optional rFirst As Range) As Double
    Dim leftMonths As Double: leftMonths = Int(nMonths / 2)
    Dim bEven As Boolean: bEven = leftMonths = nMonths / 2
    Dim rightMonths As Integer
    If bEven Then
        rightMonths = leftMonths - 1
    Else
        rightMonths = leftMonths
    End If
    
    Dim rOffs As Range: Set rOffs = r.Offset(0, rightMonths)
    CentralSum = TnM(rOffs, nMonths, rFirst)

End Function



'>> Timestamps

Sub SetTimestampOnAllSheets()
    On Error Resume Next
    Dim ws As Worksheet
    Dim ts As Double: ts = Now()
    For Each ws In Worksheets
        ws.Range(R_TIMESTAMP).Value2 = ts
    Next
End Sub

Sub SetTimestamps(ParamArray aWorksheets())
    On Error Resume Next
    Dim ix As Integer
    Dim ws As Worksheet
    For ix = LBound(aWorksheets) To UBound(aWorksheets)
        If TypeOf aWorksheets(ix) Is Worksheet Then
            Set ws = aWorksheets(ix)
        Else
            Set ws = Worksheets(aWorksheets(ix))
        End If
        ws.Range(R_TIMESTAMP).Value2 = Now()
    Next
End Sub



'>> Util

' in Macros
' /22.05.2014/
' /22.05.2014/
' /24.05.2014/
' /28.05.2014/ dim fix
' /04.06.2014/ ReplaceSumFormulasWithValues, HideFalseRows, ChangeColumnWidthForRange, Shrink/GrowSelectedColumns
' /05.06.2014/ RemoveRowsForZerosInColumn, GetCellColor
' /19.06.2014/ ColumnLetter
' 02.07.15 baselined

Function Inc(ByRef i) As Variant
    i = i + 1
    Inc = i
End Function

Function RHS(s As String) As String
    RHS = RegMatch(s, "(= *.+)$", 1)
End Function

Public Function NameExists(sRangeName As String) As Boolean
    On Error GoTo DoesntExists
    Dim rNamed As Range: Set rNamed = Range(sRangeName)
    NameExists = True
    Exit Function
DoesntExists:
    NameExists = False
End Function

Function IndentLevel(r As Range)
    IndentLevel = r.IndentLevel
End Function

Function IndentLevelVolatile(r As Range)
    Application.Volatile
    IndentLevelVolatile = r.IndentLevel
End Function

Function IsCompleteColumn(r As Range)
    IsCompleteColumn = r.Rows.Count = Sheets(1).Rows.Count
End Function


Sub CheckIsCompleteColumn()
    MsgBox IsCompleteColumn(Selection)
End Sub


Function CountA(r As Range)
    CountA = Application.CountA(r)
End Function


Function MatchWorksheetFunction(v As Variant, r As Range)
    Dim res As Variant: res = Application.match(v, r, 0)
    If IsError(res) Then MatchWorksheetFunction = 0 Else MatchWorksheetFunction = res
End Function

Function MatchManual(v As Variant, r As Range)
    Dim rCell As Range
    Dim ix As Long
    On Error GoTo err
    For Each rCell In r
        ix = ix + 1
        If rCell.Value2 = v Then
            MatchManual = ix
            Exit Function
        End If
err:
    Next
    MatchManual = 0
End Function


Function match(v As Variant, r As Range)
    match = MatchWorksheetFunction(v, r)
End Function

Function Matched(v As Variant, r As Range) As Range
    Dim ix: ix = match(v, r)
    If ix > 0 Then Set Matched = r.Cells(ix)
End Function


Function MatchedValue(v As Variant, r As Range) As Variant
    Dim ix: ix = match(v, r)
    If ix > 0 Then MatchedValue = r.Cells(ix).Value2
End Function


Sub SortColumns(rSortRange As Range)
    Application.ScreenUpdating = False
    
    Dim wsLastSheet As Worksheet: Set wsLastSheet = ActiveSheet
    Dim ws As Worksheet: Set ws = rSortRange.Worksheet
    
    ws.Activate
    rSortRange.Select
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ActiveCell.Range _
        ("A1:A152"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ws.Sort.SortFields.Add key:=ActiveCell. _
        Offset(0, 1).Range("A1:A152"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ws.Sort
        .SetRange Selection
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ws.Cells(1, 1).Select
    
    wsLastSheet.Activate
    Application.ScreenUpdating = True
End Sub



Sub AssignTextFormatToSelection()
    ' http://stackoverflow.com/questions/10801537/unable-to-set-the-numberformat-property-of-the-range-class
    
    Application.ScreenUpdating = False
    
    Dim bOk As Boolean: bOk = False
    Dim ix As Long
    For ix = 1 To Selection.Cells.Count
        Dim r As Range: Set r = Selection.Cells(ix)
        r.NumberFormat = "@"
        Dim sOldValue As String: sOldValue = r.value
        r.FormulaR1C1 = ""
        r.FormulaR1C1 = sOldValue
    Next ix

    Application.ScreenUpdating = True
End Sub


Sub AssignStandardFormatToSelection()
    Application.ScreenUpdating = False
    
    Dim bOk As Boolean: bOk = False
    Dim ix As Long
    For ix = 1 To Selection.Cells.Count
        Dim r As Range: Set r = Selection.Cells(ix)
        r.NumberFormat = "General"
        Dim vOldValue As Variant: vOldValue = r.value
        r.FormulaR1C1 = ""
        r.FormulaR1C1 = vOldValue
    Next ix

    Application.ScreenUpdating = True
End Sub


Sub ReplaceSumFormulasWithValues()
    Dim rCell As Range
    For Each rCell In Selection
        If Left(rCell.formula, 5) = "=SUM(" Then rCell.formula = rCell.value
    Next
    Application.StatusBar = False
End Sub


Sub HideFalseRows()
    Application.ScreenUpdating = False
    Dim rTrueFalse As Range: Set rTrueFalse = Selection
    Dim rTF As Range
    For Each rTF In rTrueFalse
        Dim bHide As Boolean: bHide = Not rTF
        If bHide Then
            rTF.EntireRow.Hidden = True
        End If
    Next
    Application.ScreenUpdating = True
End Sub


Sub ChangeColumnWidthForRange(rFYtoFY As Range, dWidth As Double)
    Dim rToChange As Range
    Dim nToChange As Integer: nToChange = rFYtoFY.Columns.Count
    Dim ixToChange As Integer: ixToChange = 1
    For Each rToChange In rFYtoFY
        Dim bFirstFY As Boolean: bFirstFY = ixToChange = 1
        Dim bLastFY As Boolean: bLastFY = ixToChange = nToChange
        If Not (bFirstFY Or bLastFY) Then
            rToChange.EntireColumn.ColumnWidth = dWidth
        End If
        ixToChange = ixToChange + 1
    Next
End Sub

Sub ShrinkSelectedColumns()
    ChangeColumnWidthForRange Selection, 0.1
End Sub

Sub GrowSelectedColumns()
    ChangeColumnWidthForRange Selection, 5.33
End Sub


Function ColumnLetter(lngCol As Long) As String
    ' http://stackoverflow.com/questions/12796973/vba-function-to-convert-column-number-to-letter
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    ColumnLetter = vArr(0)
End Function


Function RangesAreEqual(r1 As Range, r2 As Range) As Variant
    If r1.Columns.Count <> r2.Columns.Count Or r1.Rows.Count <> r2.Rows.Count Then
        RangesAreEqual = False
    Else
        Dim ixRow As Integer
        Dim ixCol As Integer
        For ixRow = 1 To r1.Rows.Count
            For ixCol = 1 To r1.Columns.Count
                Dim v1: v1 = r1.Cells(ixRow, ixCol).Value2
                Dim v2: v2 = r2.Cells(ixRow, ixCol).Value2
                If v1 <> v2 Then
                    RangesAreEqual = False
                    Exit Function
                End If
            Next
        Next
    End If
    RangesAreEqual = True
End Function



'>> Util2

' ((KBPXJTY)) merge Unique() and UniqueValues(), see also Sorted...

' 02.07.15 baselined
' 27.07.15 new IsRow/ColumnVector
' 07.09.15 copied Util2 from Vertragsdatenbank
' 01.02.16 ConvertLocale()

'*************************************
' special sums
'*************************************

Function SafeAverage(rValues As Range) As Double
    Dim dSum As Double
    Dim n As Integer
    Dim rValue As Range
    For Each rValue In rValues
        If IsEmpty(rValue) Or IsError(rValue) Or Not IsNumeric(rValue) Then
            'skip
        Else
            dSum = dSum + AsDouble(rValue.Value2)
            n = n + 1
        End If
    Next
    SafeAverage = dSum / n
End Function


Function SumSafe(v, Optional bAbs As Boolean) As Double
    ' only works for 1-dim and 2-dim arrays and ranges
    
    Dim a()
    If TypeOf v Is Range Then
        Dim r As Range: Set r = v
        a = RangeToArray(r)
    Else
        If IsArray(v) Then
            a = v
        Else
            On Error Resume Next
            SumSafe = v
            Exit Function
        End If
    End If
    
    SumSafe = 0
    Dim n1 As Long: n1 = ArrayCount(a, 1)
    Dim n2 As Long: n2 = ArrayCount(a, 2)
        
    On Error Resume Next
        
    If n2 = 0 Then
        ' 1-dimensional
        Dim ix As Long
        For ix = LBound(a) To n1
            If bAbs Then
                SumSafe = SumSafe + Abs(a(ix))
            Else
                SumSafe = SumSafe + a(ix)
            End If
        Next
    Else
    
        ' 2-dimensional (maybe coming from a range)
        Dim ix1 As Long: Dim ix2 As Long
        For ix1 = LBound(a, 1) To n1
            For ix2 = LBound(a, 2) To n2
                If bAbs Then
                    SumSafe = SumSafe + Abs(a(ix1, ix2))
                Else
                    SumSafe = SumSafe + a(ix1, ix2)
                End If
            Next
        Next
    End If
End Function


Function SumVisible(r As Range) As Double
    ' not needed, call TEILERGEBNIS (SUBTOTAL) w/ function numbers > 100 ("ignores hidden values")
    ' https://support.office.com/en-us/article/SUBTOTAL-function-7b027003-f060-4ade-9040-e478765b9939
    Dim rRow As Range
    Dim rCell As Range
    On Error Resume Next
    For Each rRow In r.Rows
        If Not rRow.Hidden Then
            For Each rCell In rRow.Cells
                SumVisible = SumVisible + rCell.Value2
            Next
        End If
    Next
End Function


Function SafeSum(v, Optional bAbs As Boolean = False) As Double
    SafeSum = SumSafe(v, bAbs)
End Function

Function SumAbs(v) As Double
    SumAbs = SafeSum(v, True)
End Function

Function AbsSum(v) As Double
    AbsSum = SafeSum(v, True)
End Function


Function SafeAdd(ParamArray v()) As Double
    Dim ix
    Dim dSum As Double
    For ix = LBound(v) To UBound(v)
        Dim d As Double: d = 0
        On Error Resume Next
        d = v(ix)
        On Error GoTo 0
        dSum = dSum + d
    Next
    SafeAdd = dSum
End Function


'*************************************
' strings
'*************************************

Function CountLeadingSpaces(s As String)
    CountLeadingSpaces = Len(s) - Len(LTrim(s))
End Function


'*************************************
' math
'*************************************

Function Ceiling(d As Double) As Double
    Ceiling = Application.WorksheetFunction.Ceiling(d, 1)
End Function

Function Sign(d As Double)
    If d = 0 Then
        Sign = 0
    Else
        If d > 0 Then Sign = 1 Else Sign = -1
    End If
End Function


Function Min(v1, v2) As Variant
    If v1 < v2 Then Min = v1 Else Min = v2
End Function

Function Max(v1, v2) As Variant
    If v1 > v2 Then Max = v1 Else Max = v2
End Function


'Function Equals(d1 As Double, d2 As Double, Optional dEpsilon As Double) As Boolean
'    If dEpsilon = 0 Then dEpsilon = EPSILON
'    Equals = Abs(d1 - d2) < dEpsilon
'End Function

Function AllEquals(dEpsilon As Double, ParamArray v()) As Boolean
    Dim bEqual As Boolean: bEqual = True
    Dim dFirst As Double
    Dim ixParam As Integer
    For ixParam = LBound(v) To UBound(v)
        Dim d As Double: d = v(ixParam)
        If ixParam = LBound(v) Then
            dFirst = d
        Else
            bEqual = Equals(dFirst, d, dEpsilon)
        End If
    Next
    AllEquals = bEqual
End Function

Function Equals(v1, v2, Optional dEpsilon As Double) As Boolean
    Dim d1 As Double, d2 As Double
    On Error GoTo err
    
    If TypeOf v1 Is Range Then
        Dim r1 As Range: Set r1 = v1
        If TypeOf v2 Is Range Then
            Dim r2 As Range: Set r2 = v2
            If r1.Cells.Count <> r2.Cells.Count Then
                Equals = False
            Else
                Dim ixCell As Integer
                For ixCell = 1 To r1.Cells.Count
                    d1 = r1.Cells(ixCell)
                    d2 = r2.Cells(ixCell)
                    If Not Equals(d1, d2, dEpsilon) Then
                        Equals = False
                        Exit Function
                    End If
                Next
                Equals = True
            End If
        Else
            If r1.Cells.Count <> 1 Then
                Equals = False
            Else
                d1 = r1.Cells(1)
                Equals = Equals(d1, v2, dEpsilon)
            End If
        End If
    Else
        d1 = v1
        d2 = v2
        If dEpsilon = 0 Then dEpsilon = EPSILON
        Equals = Abs(d1 - d2) < dEpsilon
    End If
    Exit Function
    
err:
    Equals = False
End Function



Function AsDouble(v) As Double
    If IsArray(v) Then
        On Error Resume Next
        AsDouble = v(LBound(v, 1), LBound(v, 2))
    Else
        If IsDate(v) Then
            AsDouble = DateSerial(Year(v), Month(v), Day(v))
        Else
            If IsNumeric(v) Then
                AsDouble = v
            Else
                AsDouble = 0
            End If
        End If
    End If
End Function


Function AsBoolean(v) As Boolean
    On Error Resume Next
    AsBoolean = v
End Function


Sub MakeNumeric()
    SetSilentApplicationState
    On Error GoTo done
    
    Dim rSelection As Range: Set rSelection = Selection.SpecialCells(xlCellTypeConstants)
    Dim r As Range
    For Each r In rSelection
        If Not (IsEmpty(r)) Then
            If RegMatch(r.Value2, "^ *[0-9.,()+-]+$") Then
                r = AsDouble(r)
            End If
        End If
    Next
    
done:
    RevertApplicationState
End Sub


'*************************************
' tests "Is..."
'*************************************

Function IsAllCellsEmpty(rAll As Range)
    Dim r As Range
    For Each r In rAll
        If Not IsEmpty(rAll) Then
            IsAllCellsEmpty = False
            Exit Function
        End If
    Next
    IsAllCellsEmpty = True
End Function


'*************************************
' tests "If..."
'*************************************

Function IfZero(v, vIfZero, Optional vIfNotZero)
    If v = 0 Then
        IfZero = vIfZero
    Else
        If IsMissing(vIfNotZero) Then
            IfZero = v
        Else
            IfZero = vIfNotZero
        End If
    End If
End Function


Function IfNothing(v, onNothing)
    IfNothing = onNothing
    If IsEmpty(v) Then Exit Function
    Dim r As Range: Set r = v
    If IsEmpty(r) Then Exit Function
    IfNothing = r.Value2
End Function


Function IfEmpty(v, vIfEmpty, Optional vIfNotEmpty)
    If IsEmpty(v) Then
        IfEmpty = vIfEmpty
    Else
        If IsMissing(vIfNotEmpty) Then
            IfEmpty = v
        Else
            IfEmpty = vIfNotEmpty
        End If
    End If
End Function


'*************************************
' decimal point = point
'*************************************

Function ConvertLocale(s As String) As String
    s = Replace(s, ".", "~")
    s = Replace(s, ",", ".")
    ConvertLocale = Replace(s, "~", ".")
End Function


'*************************************
' PL growth rate / percent of
'*************************************

Function RepeatString(s As String, n As Integer) As String
    Dim ix As Integer
    For ix = 1 To n
        RepeatString = RepeatString + s
    Next ix
End Function

Function VerboseGrowthRate(TY As Double, LY As Double, Optional dMaximum As Double, Optional dMinimum As Double, Optional nDigits As Integer, Optional bOtherMinus As Boolean) As Variant
    ' dMaximum: $H$2, dMinimum = 1%
    ' TY: F9, LY: C9
    ' =WENNFEHLER(WENN(ODER(ABS(F9/ABS(C9)-VORZEICHEN(C9))<1%;ABS(F9/ABS(C9)-VORZEICHEN(C9))>$H$2);"";F9/ABS(C9)-VORZEICHEN(C9));"")
    
    VerboseGrowthRate = ""
    Dim sMinus As String
    ' &H336: Verbindungszeichen langer Querstrich (Overlay)
    ' &H2013: Bindestrich
    If bOtherMinus Then sMinus = ChrW(&H2013) Else sMinus = "-"
    If dMaximum = 0 Then dMaximum = PERCENTAGE_TOO_BIG
    If LY = 0 Then
        ' avoid division by zero
    Else
        If Abs(TY - LY) < EPSILON Then
            ' equal, don't show "0%"
        Else
            Dim growth As Double: growth = TY / Abs(LY) - Sign(LY)
            If Abs(Abs(growth) - 1) < EPSILON And (TY = 0 Or LY = 0) Then
                ' don't show "100%" or "-100%"
            Else
                If Abs(growth) + EPSILON > dMaximum Then
                    ' suppress percentages which are too big
                Else
                    If Abs(growth) + EPSILON < dMinimum Then
                        ' suppress percentages which are too small
                    Else
                        VerboseGrowthRate = growth
                        'If nDigits = 0 Then
                        '    VerboseGrowthRate = ConvertLocale(Format(growth, "+0%;" & sMinus & "0%"))
                        'Else
                        '    Dim sDigits As String
                        '    sDigits = "." & RepeatString("0", nDigits)
                        '    VerboseGrowthRate = ConvertLocale(Format(growth, "+0" & sDigits & "%;" & sMinus & "0" & sDigits & "%"))
                        'End If
                    End If
                End If
            End If
        End If
    End If
End Function


Function ev(v)
    ev = Evaluate(v & "+0")
End Function

Function f(v, Optional sFormat As String) As String
    If sFormat <> "" Then v = Format(v, sFormat)
    ' &H2013: Bindestrich
    v = Replace(v, "-", ChrW(&H2013))
    f = v
End Function


Function SS(ParamArray a() As Variant) As Double
    Dim dSum As Double: dSum = 0
    Dim v
    For Each v In a
        If TypeName(v) = "Range" Then
            Dim r As Range: Set r = v
            Dim vCell
            For Each vCell In r.Cells
                dSum = dSum + f(vCell)
            Next
        Else
            On Error Resume Next
            dSum = dSum + f(v)
            On Error GoTo 0
        End If
    Next
    SS = dSum
End Function


Function ToNumber(v)
    If IsNumeric(v) Then
        ToNumber = v
    Else
        On Error GoTo err
        v = Replace(v, "â€“", "-")
        v = Replace(v, ".", "#")
        v = Replace(v, ",", ".")
        v = Replace(v, "#", ",")
        ToNumber = Evaluate(v & "+0")
        Exit Function
err:
        ToNumber = 0
    End If
 End Function


Function bla(v)
    If IsNumeric(v) Then
        bla = v
    Else
        On Error GoTo err
        v = Replace(v, "â€“", "-")
        v = Replace(v, ".", "#")
        v = Replace(v, ",", ".")
        v = Replace(v, "#", ",")
        bla = Evaluate(v & "+0")
        Exit Function
err:
        bla = 0
    End If
 End Function


Function VerbosePercentOf(d As Double, dTotal As Double, Optional dMinimum As Double, Optional bShow100 As Boolean, Optional bOtherMinus As Boolean) As Variant
    ' C9: d, C$9: dTotal, D2: dMinimum, bShow100 = True
    ' =WENN(ODER(C9=0;C$9=0;ABS(C9/C$9)<$D$2);"";C9/C$9)
    
    VerbosePercentOf = ""
    Dim sMinus As String
    ' &H336: Verbindungszeichen langer Querstrich (Overlay)
    ' &H2013: Bindestrich (–, - ist normales Minus)
    If bOtherMinus Then sMinus = ChrW(&H2013) Else sMinus = "-"
    If dTotal = 0 Then
        ' avoid division by zero
    Else
        If d = 0 Then
            ' don't show "0,0%"
        Else
            If Abs(d - dTotal) < EPSILON And Not bShow100 Then
                ' equal, don't show "100,0%"
            Else
                Dim dfrac As Double: dfrac = d / dTotal
                If Abs(dfrac) + EPSILON < dMinimum Then
                    ' don't show percentage which are too small
                Else
                    'VerbosePercentOf = ConvertLocale(Format(dfrac, "0.0%;" & sMinus & "0.0%"))
                    VerbosePercentOf = dfrac
                End If
            End If
        End If
    End If
End Function


Function VerbosePercentDifference(dTY, dTotalTY, dLY, dTotalLY, Optional dMinimum As Double, Optional dMaximum As Double, Optional bRevertSign As Boolean, Optional bOtherMinus As Boolean) As Variant
    ' F9: dTY, F$9: dTotalTY, C9: dLY, C$9: dTotalLY, dMinimum = 1%, dMaximum = $H$2
    ' =WENNFEHLER(WENN(ODER(ABS(F9/F$9-C9/C$9)<0,1%;ABS(F9/F$9-C9/C$9)>$J$2);"";F9/F$9-C9/C$9);"")

    VerbosePercentDifference = ""
    Dim sMinus As String
    ' &H336: Verbindungszeichen langer Querstrich (Overlay)
    ' &H2013: Bindestrich
    If bOtherMinus Then sMinus = ChrW(&H2013) Else sMinus = "-"
    
    If dMaximum = 0 Then dMaximum = 10000000000#
    If dTotalTY = 0 Or dTotalLY = 0 Then
        ' either % is N/A
    Else
        On Error Resume Next
        Dim dFracTY As Double: dFracTY = dTY / dTotalTY
        Dim dFracLY As Double: dFracLY = dLY / dTotalLY
        On Error GoTo 0
        Dim dDiff As Double: dDiff = dFracTY - dFracLY
        If bRevertSign Then dDiff = -dDiff
        If Abs(dDiff) < 0.000001 Then
        Else
            If Abs(dDiff) < dMinimum Or Abs(dDiff) > dMaximum Then
                ' don't show percentage which are too small
            Else
                'VerbosePercentDifference = ConvertLocale(Format(dDiff, "+0.0%;" & sMinus & "0.0%"))
                VerbosePercentDifference = dDiff
            End If
        End If
    End If
End Function


Function SafeGrowthRate(dt As Double, dL As Double) As Variant
    '=WENN(ISTFEHLER(C27/C15);"";WENN(C27/C15-1>10;"> +1000%";WENN(C27/C15-1<-10;"< -1000%";C27/C15-1)))
    Dim dGrowth As Double
    On Error GoTo err
    dGrowth = dt / dL - 1
    If dGrowth > 10 Then
        SafeGrowthRate = "> +1000%"
    Else
        If dGrowth < -10 Then
            SafeGrowthRate = "< -1000%"
        Else
            If dGrowth > 1 Then
                'SafeGrowthRate = Format(dGrowth, "+0%;-0%")
                SafeGrowthRate = dGrowth
            Else
                If dGrowth < -1 Then
                    'SafeGrowthRate = Format(dGrowth, "+0%;-0%")
                    SafeGrowthRate = dGrowth
                Else
                    'SafeGrowthRate = Format(dGrowth, "+0.0%;-0.0%")
                    SafeGrowthRate = dGrowth
                End If
            End If
        End If
    End If
    Exit Function
err:
    SafeGrowthRate = ""
End Function


Function SafeShare(dDividend, dDivisor, Optional bShowZero As Boolean) As Variant
    Dim dShare As Double
    If Not bShowZero And dDividend = 0 Then GoTo err
    On Error GoTo err
    dShare = dDividend / dDivisor
    If dShare > 10 Then
        SafeShare = "> 1000%"
    Else
        If dShare < -10 Then
            SafeShare = "< 1000%"
        Else
            If dShare > 1 Then
                'SafeShare = Format(dShare, "0%")
                SafeShare = dShare
            Else
                If dShare < -1 Then
                    'SafeShare = Format(dShare, "0%")
                    SafeShare = dShare
                Else
                    'SafeShare = Format(dShare, "0.0%")
                    SafeShare = dShare
                End If
            End If
        End If
    End If
    Exit Function
err:
    SafeShare = ""
End Function


'*************************************
' arrays / vectors
'*************************************

Function IsArrayAllocated(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = IsArray(arr) And LBound(arr, 1) <= UBound(arr, 1)
End Function
                           

Function IsColumnRange(r As Range) As Boolean
    IsColumnRange = r.Columns.Count = 1
End Function

Function IsRowRange(r As Range) As Boolean
    IsRowRange = r.Rows.Count = 1
End Function

Function IsVectorRange(r As Range) As Boolean
    IsVectorRange = IsColumnRange(r) Or IsRowRange(r)
End Function

Function IsVector(v() As Variant, Optional bStrict As Boolean) As Boolean
    Dim n1 As Long: n1 = ArrayCount(v, 1)
    Dim n2 As Long: n2 = ArrayCount(v, 2)
    
    IsVector = False
    
    If n1 = 0 And n2 = 0 Then Exit Function
    If n2 > 0 And bStrict Then Exit Function
    If n1 > 1 And n2 > 1 Then Exit Function
    
    IsVector = True
End Function


Function IsRowVector(v() As Variant) As Boolean
    Dim n1 As Long: n1 = ArrayCount(v, 1)
    Dim n2 As Long: n2 = ArrayCount(v, 2)

    IsRowVector = False
    
    If n1 = 0 And n2 = 0 Then Exit Function
    If n1 > 1 And n2 > 1 Then Exit Function
    If n1 > 1 And n2 = 1 Then Exit Function
    
    IsRowVector = True
End Function


Function IsColumnVector(v() As Variant) As Boolean
    Dim n1 As Long: n1 = ArrayCount(v, 1)
    Dim n2 As Long: n2 = ArrayCount(v, 2)

    IsColumnVector = False
    
    If n2 = 0 Then Exit Function
    If n1 = 0 And n2 = 0 Then Exit Function
    If n1 > 1 And n2 > 1 Then Exit Function
    If n1 = 1 And n2 > 1 Then Exit Function
    
    IsColumnVector = True
End Function


Function ArrayToCollection(v() As Variant) As Collection
    ' funktioniert nur fÃ¼r max 2-dimensionale arrays
    
    Set ArrayToCollection = New Collection
    
    Dim n2 As Long: n2 = ArrayCount(v, 2)
    Dim ix1 As Long, ix2 As Long
    For ix1 = LBound(v, 1) To UBound(v, 1)
        If n2 = 0 Then
            ArrayToCollection.Add v(ix1)
        Else
            For ix2 = LBound(v, 2) To UBound(v, 2)
                ArrayToCollection.Add v(ix1, ix2)
            Next ix2
        End If
    Next ix1
End Function


Function StringArrayToCollection(v()) As Collection
    ' 12.02.22 was v() as string
    ' funktioniert nur fÃ¼r max 2-dimensionale arrays
    
    Set StringArrayToCollection = New Collection
    
    On Error Resume Next
    Dim n1 As Long: n1 = UBound(v, 1) - LBound(v, 1) + 1
    Dim n2 As Long: n2 = UBound(v, 2) - LBound(v, 2) + 1
    On Error GoTo 0
    
    Dim ix1 As Long
    Dim ix2 As Long
    For ix1 = LBound(v, 1) To n1 - 1
        If n2 = 0 Then
            StringArrayToCollection.Add v(ix1)
        Else
            For ix2 = LBound(v, 2) To n2 - 1
                StringArrayToCollection.Add v(ix1, ix2)
            Next ix2
        End If
    Next ix1
End Function


' 12.02.22 added
Public Function StringsToCollection(ParamArray pasStrings()) As Collection
    Dim asStrings(): asStrings = ParamArrayDelegated(pasStrings)
    Set StringsToCollection = StringArrayToCollection(asStrings)
End Function


Sub MakeVector(ByRef v() As Variant)
    If IsVector(v, True) Then
        ' do nothing, is a vector
    Else
    
        v = CollectionToArray(ArrayToCollection(v))
    
    End If
End Sub


Function VectorCount(v() As Variant) As Long
    VectorCount = ArrayCount(v)
End Function

Function ArrayCount(a() As Variant, Optional ixDimension As Integer) As Long
    On Error Resume Next
    If ixDimension = 0 Then ArrayCount = UBound(a) - LBound(a) + 1 Else ArrayCount = UBound(a, ixDimension) - LBound(a, ixDimension) + 1
End Function


Public Function GetPartialCollection(c As Collection, ixFirst As Long, ixLast As Long) As Collection
    Debug.Assert ixLast >= ixFirst
    Dim cPartial As Collection: Set cPartial = New Collection
    Dim ix As Long
    For ix = ixFirst To ixLast
        cPartial.Add c(ix)
    Next ix
    Set GetPartialCollection = cPartial
End Function


Public Function RangeToArray(r As Range) As Variant()
    ' simple assignment works
    If r.Cells.Count = 1 Then
        Dim a(): ReDim a(1, 1)
        a(1, 1) = r.Value2
    Else
        RangeToArray = r
    End If
End Function


Public Function CreateVector(ParamArray v() As Variant) As Variant()
    Dim n As Long: n = 0
    Dim vCell As Variant
    Dim r As Range
    For Each vCell In v
        If TypeOf vCell Is Range Then
            Set r = vCell
            n = n + r.Cells.Count
        Else
            n = n + 1
        End If
    Next
    
    Dim a(): ReDim a(1 To n)
    Dim ix As Long: ix = 0
    
    For Each vCell In v
        If TypeOf vCell Is Range Then
            Set r = vCell
            Dim rCell As Range
            For Each rCell In r.Cells
                ix = ix + 1
                a(ix) = rCell.Value2
            Next
        Else
            ix = ix + 1
            a(ix) = vCell
        End If
    Next
    
    CreateVector = a
End Function
    
Public Function GetRowVector(a() As Variant, ixRow As Long) As Variant()
    Dim nCols As Long: nCols = ArrayCount(a, 2)
    Dim v(): ReDim v(LBound(a, 1) To UBound(a, 1))
    Dim ixCol As Long
    For ixCol = LBound(a, 1) To UBound(a, 1)
        v(ixCol) = a(ixRow, ixCol)
    Next ixCol
    GetRowVector = v
End Function

Public Function GetColVector(a() As Variant, ixCol As Long) As Variant()
    Dim nRows As Long: nRows = ArrayCount(a, 1)
    Dim v(): ReDim v(LBound(a, 1) To UBound(a, 1))
    Dim ixRow As Long
    For ixRow = LBound(a, 1) To UBound(a, 1)
        v(ixRow) = a(ixRow, ixCol)
    Next ixRow
    GetColVector = v
End Function

Public Function GetColumnVector(a() As Variant, ixCol As Long) As Variant()
    GetColumnVector = GetColVector(a, ixCol)
End Function



Public Function ArrayToVector(a() As Variant) As Variant()
    Dim nRows As Long: nRows = ArrayCount(a, 1)
    Dim nCols As Long: nCols = ArrayCount(a, 2)
    
    Dim ixElement As Long: ixElement = 1
    Dim nElements As Long
    Dim ixRow As Long
    Dim aVector()
    
    If nCols = 0 Then
    
        nElements = nRows
        ReDim aVector(1 To nElements)
        For ixRow = 1 To nRows
            aVector(ixElement) = a(ixRow)
            ixElement = ixElement + 1
        Next
    
    Else
    
        nElements = nRows * nCols
        ReDim aVector(1 To nElements)
        Dim ixCol As Long
        For ixRow = 1 To nRows
            For ixCol = 1 To nCols
                aVector(ixElement) = a(ixRow, ixCol)
                ixElement = ixElement + 1
            Next
        Next
    End If
    
    ArrayToVector = aVector
End Function


Public Function GetPartialArray(a() As Variant, ixFirstRow As Long, ixLastRow As Long) As Variant()
    Debug.Assert ixLastRow >= ixFirstRow
    Dim nCols As Long: nCols = ArrayCount(a, 2)
    Dim nRows As Long: nRows = ixLastRow - ixFirstRow + 1
    Dim aRows(): ReDim aRows(1 To nRows, 1 To nCols)
    Dim ixRow As Long
    For ixRow = 1 To nRows
        Dim ixCol As Long
        For ixCol = 1 To nCols
            aRows(ixRow, ixCol) = a(ixFirstRow + ixRow - 1, ixCol)
        Next ixCol
    Next ixRow
    GetPartialArray = aRows
End Function


Public Function GetPartialVector(v() As Variant, ixFirst As Long, ixLast As Long) As Variant()
    Debug.Assert ixLast >= ixFirst
    Dim nElements As Long: nElements = ixLast - ixFirst + 1
    Dim vPartial(): ReDim vPartial(1 To nElements)
    Dim ix As Long
    For ix = 1 To nElements
        vPartial(ix) = v(ixFirst + ix - 1)
    Next ix
    GetPartialVector = vPartial
End Function


Sub AddToVector(ByRef vTarget() As Variant, ByRef vToSum() As Variant)
    Dim ix As Long
    
    ' Probably possible to do some magic on differently based vectors, but caller should make sure they are compatible.
    ' Usually 1-based anyways.
    Debug.Assert LBound(vTarget) = LBound(vToSum)
    
    ' Necessary because it might be impossible to convert a Variant to Double.
    On Error Resume Next
    
    Dim nElementsToSum As Long: nElementsToSum = Min(VectorCount(vTarget), VectorCount(vToSum))
    For ix = LBound(vTarget, 1) To nElementsToSum
        Dim dToSum As Double
        dToSum = vToSum(ix)
        vTarget(ix) = vTarget(ix) + dToSum
    Next ix
End Sub


Function VectorSum(v() As Variant) As Double
    Dim sum As Double
    Dim ix As Long
    For ix = LBound(v, 1) To UBound(v, 1)
        sum = sum + v(ix)
    Next ix
    VectorSum = sum
End Function


Function VectorAbsSum(v() As Variant) As Double
    Dim sum As Double
    Dim ix As Long
    On Error Resume Next
    For ix = LBound(v, 1) To UBound(v, 1)
        Dim dAbsToSum As Double
        dAbsToSum = Abs(v(ix))
        sum = sum + dAbsToSum
    Next ix
    VectorAbsSum = sum
End Function


Function CreateCollection(ParamArray v() As Variant) As Collection
    Dim c As Collection: Set c = New Collection
    Dim ixElement As Long
    For ixElement = LBound(v) To UBound(v)
        c.Add v(ixElement)
    Next ixElement
    Set CreateCollection = c
End Function


Function CC(ParamArray v() As Variant) As Collection
    Set CC = CreateCollection(v)
End Function


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


Function AutoOrientVector(av() As Variant) As Variant()
    Debug.Assert IsVector(av)
    
    On Error Resume Next
    Dim rCaller As Range: Set rCaller = Application.Caller
    On Error GoTo 0
    
    If rCaller Is Nothing Then
        AutoOrientVector = av
    Else
    
        If IsColumnRange(rCaller) Then
            If IsColumnVector(av) Then
                AutoOrientVector = av
            Else
                AutoOrientVector = Application.WorksheetFunction.Transpose(av)
            End If
        Else
            If IsRowVector(av) Then
                AutoOrientVector = av
            Else
                AutoOrientVector = Application.WorksheetFunction.Transpose(av)
            End If
        End If
    End If
End Function


Function Prune(a() As Variant) As Variant()
    Dim nPassedRows As Long: nPassedRows = ArrayCount(a, 1)
    Dim nPassedCols As Long: nPassedCols = ArrayCount(a, 2)
    Dim bIsOneDimensional As Boolean: bIsOneDimensional = nPassedCols = 0
    
    Dim rCaller As Range: Set rCaller = Application.Caller
    Dim nCallerRows As Long: nCallerRows = rCaller.Rows.Count
    Dim nCallerCols As Long: nCallerCols = rCaller.Columns.Count
    
    Dim aPruned(): ReDim aPruned(1 To nCallerRows, 1 To nCallerCols)
    
    Dim ixRow As Long
    Dim ixCol As Long
    For ixRow = 1 To nCallerRows
        For ixCol = 1 To nCallerCols
            Dim bInsidePassed As Boolean: bInsidePassed = ixRow <= nPassedRows And ((bIsOneDimensional And ixCol = 1) Or (ixCol <= nPassedCols))
            If bInsidePassed Then
                If bIsOneDimensional Then
                    aPruned(ixRow, ixCol) = a(ixRow)
                Else
                    aPruned(ixRow, ixCol) = a(ixRow, ixCol)
                End If
            Else
                aPruned(ixRow, ixCol) = "."
            End If
        Next ixCol
    Next ixRow
    
    Prune = aPruned
End Function


Function UniqueValues(xx As Variant) As Variant()
    Dim av()
    If TypeOf xx Is Range Then
        Dim r As Range: Set r = xx
        av = RangeToVector(r)
    Else
        av = xx
    End If
    Dim dict As New Dictionary
    Dim ix As Integer
    For ix = LBound(av) To UBound(av)
        Dim v: v = av(ix)
        If Not dict.Exists(v) Then dict.Add v, v
    Next
    Dim avUnique(): ReDim avUnique(1 To dict.Count)
    For ix = 1 To dict.Count
        avUnique(ix) = dict.Keys(ix - 1)
    Next ix
    UniqueValues = AutoOrientVector(avUnique)
End Function


' ((KBPXJTY)) merge Unique() and UniqueValues(), see also Sorted...
Function Unique(c As Collection) As Collection
    Set Unique = New Collection
    Dim dict As New Dictionary
    Dim v
    For Each v In c
        If dict.Exists(v) Then
            ' ignore
        Else
            dict.Add v, v
            Unique.Add v
        End If
    Next
End Function


Function ExpandedRange(r As Range) As Range
    Dim ixRow As Integer
    Do While True
        If IsEmpty(r.Offset(ixRow + 1, 0)) Then Exit Do
        ixRow = ixRow + 1
    Loop
    
    Dim ixCol As Integer
    Do While True
        If IsEmpty(r.Offset(0, ixCol + 1)) Then Exit Do
        ixCol = ixCol + 1
    Loop
    Set ExpandedRange = r.Worksheet.Range(r, r.Offset(ixRow, ixCol))
End Function


'*************************************
' dictionaries
'*************************************

Function DictionaryKeysToCollection(d As Dictionary) As Collection
    Dim c As New Collection
    Dim v
    For Each v In d.Keys
        c.Add v
    Next
    Set DictionaryKeysToCollection = c
End Function


Function MergeDictionaryKeys(dict1 As Dictionary, dict2 As Dictionary) As Collection
    Dim dictMerged As New Dictionary
    Dim vKey
    For Each vKey In dict1.Keys()
        dictMerged(vKey) = vKey
    Next
    For Each vKey In dict1.Keys()
        dictMerged(vKey) = vKey
    Next
    
    Set MergeDictionaryKeys = DictionaryKeysToCollection(dictMerged)
End Function
    
    
'*************************************
' save to sheet
'*************************************

Sub SaveCollectionToSheet(c As Collection, rAnchor As Range, Optional bNumbersWherePossible As Boolean, Optional bHorizontal As Boolean)
    Dim ix As Long
    For ix = 1 To c.Count
        Dim s As String: s = c(ix)
        Dim nSpaces As Integer: nSpaces = CountLeadingSpaces(s)
        Dim rOutput As Range
        If bHorizontal Then
            Set rOutput = rAnchor.Offset(0, ix - 1)
        Else
            Set rOutput = rAnchor.Offset(ix - 1)
        End If
        If bNumbersWherePossible Then
            rOutput.value = LTrim(s)
        Else
            rOutput.value = "'" & LTrim(s)
        End If
        If nSpaces > 0 Then rOutput.IndentLevel = nSpaces
    Next ix
End Sub


Sub SaveArrayToSheet(a() As Variant, rAnchor As Range)
    Dim nRows As Long: nRows = ArrayCount(a, 1)
    Dim nCols As Long: nCols = ArrayCount(a, 2)
    Dim rTarget As Range: Set rTarget = rAnchor.Worksheet.Range(rAnchor, rAnchor.Offset(nRows - 1, nCols - 1))
    rTarget = a
End Sub


'*************************************
' range conversion
'*************************************

Public Function RangeCollectionToAddresses(c As Collection, Optional bWithSheet As Boolean = True) As Collection
    Dim cAddresses As Collection: Set cAddresses = New Collection
    Dim r As Range
    For Each r In c
        cAddresses.Add GetAddress(r, bWithSheet)
    Next
    Set RangeCollectionToAddresses = cAddresses
End Function


Public Function RangeToCollection(r As Range, Optional bConvertToString As Boolean) As Collection
    Dim c As Collection: Set c = New Collection
    Dim ixCell As Long
    For ixCell = 1 To r.Cells.Count
        Dim v As Variant: v = r.Cells(ixCell)
        If bConvertToString Then
            Dim s As String: s = v
            c.Add s
        Else
            c.Add v
        End If
    Next ixCell
    Set RangeToCollection = c
End Function


Public Function RangeToVector(r As Range, Optional iLBound As Long = 1) As Variant()
    ' simple assignment doesn't work
    ' see http://stackoverflow.com/questions/19038697/excel-vba-populate-array-from-named-range
    Dim v()
    Dim iUBound As Long: iUBound = r.Cells.Count - (iLBound - 1)
    ReDim v(iLBound To iUBound)
    Dim ix As Long
    For ix = 1 To r.Cells.Count
        v(ix + iLBound - 1) = r.Cells(ix)
    Next ix
    RangeToVector = v
End Function


'*************************************
' sheet helpers
'*************************************

Function GetLastNonEmptyCellInColumn(rTopCellToStartSearchFrom As Range)
    'Set GetLastNonEmptyCellInColumn = rTopCellToStartFrom.End(xlDown)
    
    'alternatively, see http://www.mrexcel.com/forum/excel-questions/697125-visual-basic-applications-find-cell-below-last-row-insert-sum.html
    Dim sColumnLetter As String: sColumnLetter = ColumnLetter(rTopCellToStartSearchFrom.Column)
    With rTopCellToStartSearchFrom.Worksheet
        Set GetLastNonEmptyCellInColumn = .Range(sColumnLetter & .Rows.Count).End(xlUp)
    End With
End Function


Sub ShowAddressOfLastNonEmptyCellInColumnOfSelectedCell()
    MsgBox GetLastNonEmptyCellInColumn(Selection).Address
End Sub


Sub GotoLastInColumn()
    GetLastNonEmptyCellInColumn(Selection).Select
End Sub


Function GetPrevious(r As Range) As Range
    If IsEmpty(r) Then
        Set GetPrevious = r.End(xlUp)
    Else
        Set GetPrevious = r
    End If
End Function


' >> Versions

'*************************************
' sheet helpers
'*************************************

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


