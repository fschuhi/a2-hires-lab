Attribute VB_Name = "Modules"
Option Explicit

Private Const EXPORT_FOLDER_RELATIVE As String = "src\bas\"

Public Sub ExportModuleByName(ByVal sModuleName As String)
    Dim vbComp As Object
    Dim sExportFolder As String
    Dim sExportPath As String
    Dim sFileExtension As String
    
    sExportFolder = ThisWorkbook.Path & Application.PathSeparator & EXPORT_FOLDER_RELATIVE
    
    If Len(Dir(sExportFolder, vbDirectory)) = 0 Then
        MkDir sExportFolder
    End If
    
    Set vbComp = ThisWorkbook.VBProject.VBComponents(sModuleName)
    
    Select Case vbComp.Type
        Case 1 ' vbext_ct_StdModule
            sFileExtension = ".bas"
            
        Case 2 ' vbext_ct_ClassModule
            sFileExtension = ".cls"
            
        Case 3 ' vbext_ct_MSForm
            ' Forms need to be saved manually
            'sFileExtension = ".frm"
            Exit Sub
            
        Case 100 ' vbext_ct_Document
            sFileExtension = ".cls"
            
        Case Else
            err.Raise vbObjectError + 513, , "Unsupported component type for: " & sModuleName
            
    End Select
    
    sExportPath = sExportFolder & Application.PathSeparator & sModuleName & sFileExtension
    
    If Len(Dir(sExportPath)) > 0 Then
        Kill sExportPath
    End If
    
    vbComp.Export sExportPath
End Sub

Public Sub ExportProjectModules()
    ExportModuleByName "B_"
    ExportModuleByName "JumpStation_"
    ExportModuleByName "Macros"
    ExportModuleByName "Modules"
    ExportModuleByName "NTSCColor"
    ExportModuleByName "ScreenMemory"
    ExportModuleByName "SpriteEditor"
    ExportModuleByName "Util"
    ExportModuleByName "WorksheetsMatrix_"
End Sub

