VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormComment 
   Caption         =   "Cell Comment"
   ClientHeight    =   3451
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   6104
   OleObjectBlob   =   "UserFormComment.frx":0000
End
Attribute VB_Name = "UserFormComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public g_sOldComment As String

' defined in Macros, do not uncomment (would be initialized every Show() time with 0)
'Public g_iUserFormCommentTop As Integer
'Public g_iUserFormCommentLeft As Integer


Private Sub DoUnload()
    g_iUserFormCommentTop = Me.Top
    g_iUserFormCommentLeft = Me.Left
    Unload Me
End Sub


Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 27
            Dim r As Range: Set r = ActiveCell
            Dim sNewComment As String: sNewComment = Trim(TextBox1.Text)
            
            If sNewComment = "" Then
                If g_sOldComment <> "" Then
                    Dim res: res = MsgBox("Kommentar löschen?", vbYesNoCancel)
                    If res <> vbYes Then sNewComment = g_sOldComment
                End If
            End If
            
            r.Comment.Delete
            If sNewComment <> "" Then r.AddComment sNewComment
            DoUnload
    End Select
End Sub


Private Sub UserForm_Activate()
    g_sOldComment = ""
    On Error Resume Next
    g_sOldComment = ActiveCell.Comment.Text
    On Error GoTo 0
    TextBox1.Text = g_sOldComment
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    
    If g_iUserFormCommentTop = 0 Then g_iUserFormCommentTop = 275
    
    ' set from left side of screen
    If g_iUserFormCommentLeft = 0 Then g_iUserFormCommentLeft = 550
    
    ' alternative: set from right side of screen
    ' If g_iUserFormCommentLeft = 0 Then g_iUserFormCommentLeft = Application.Left + Application.Width - Me.Width - 50
    
    Me.Top = g_iUserFormCommentTop
    Me.Left = g_iUserFormCommentLeft
End Sub
