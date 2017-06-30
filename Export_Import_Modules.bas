Attribute VB_Name = "Export_Import_Modules"
Option Explicit

Public Sub ExportModules()
'
'*******************************************************************************
' Sub Name: ExportVBACode
'
' Sub Purpose:
' ExportModules copies all the Forms, Modules, and Classes from the active
' workbook into a folder in your Document folder called Path.
' It creates this folder if it does not exist.
' It deletes all files in the folder if it does exist.
'
' Called by: User
'
'*******************************************************************************
'

Dim bExport As Boolean
Dim wkbSource As Excel.Workbook
Dim szExportPath As String
Dim szFileName As String
Dim cmpComponent As VBIDE.VBComponent
Dim WkBkName As String
Dim Path As String

    WkBkName = GetExcelFilePathAndName("Select the Workbook Containing the VBA Code You Wish to Export", _
        "Excel Macro-Enabled Workbook(*.xlsm), *.xlsm")
    
    ''' NOTE: This workbook must be open in Excel.
    On Error Resume Next
    Set wkbSource = Application.Workbooks(GetFileName(WkBkName))
    If Err.Number <> 0 Then
        MsgBox "The workbook must be open"
        Exit Sub
    End If
    On Error GoTo 0
    
    Path = GetFolder("Select the Folder to Place the VBA Code")
    If Path = "" Then Exit Sub

    If FolderWithVBAProjectFiles(Path) = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = Path & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name
'        Debug.Print cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
'                szFileName = szFileName & ".cls"
                bExport = False
            Case vbext_ct_ActiveXDesigner
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        End If
   
    Next cmpComponent

    MsgBox "All Forms, Modules, and Classes have been exported into the " & _
        Path & " folder"
        
End Sub ' ExportModules

Public Sub ImportModules()
'
'*******************************************************************************
' Sub Name: ImportModules
'
' Sub Purpose:
' ImportModules copies all the Forms, Modules, and Classes from a folder in
' your Document folder into the active workbook. It deletes all Forms,
' Modules, and Classes from the active workbook.
'
' Called by: User
'
'*******************************************************************************
'
Dim objFSO As Scripting.FileSystemObject
Dim objFile As Scripting.File
Dim szImportPath As String
Dim cmpComponents As VBIDE.VBComponents
Dim Response As String
Dim WkBkName As String
Dim Path As String
Dim szSourceWorkbook As String
Dim wkbSource As Workbook

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = GetFileName(ActiveWorkbook.Name)
    On Error Resume Next
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    If Err.Number <> 0 Then
        MsgBox "The workbook must be open"
        Exit Sub
    End If
    On Error GoTo 0
    
    Response = MsgBox("Warning. All existing Forms, Modules, and Classes will " & _
        "be deleted from " & wkbSource.Name & ". Continue?", vbYesNo, "Warning")
    If Response = vbNo Then Exit Sub

    If WkBkName = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    Path = GetFolder("Select the Folder Containing the VBA Code You Want to Import")
    If FolderWithVBAProjectFiles(Path) = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles(Path)
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms(wkbSource.Name)

    Set cmpComponents = wkbSource.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
        
    Next objFile
    
    MsgBox "All Forms, Modules, and Classes imported"
    
End Sub ' ImportModules

Private Function FolderWithVBAProjectFiles(ByVal Path As String) As String
'
'*******************************************************************************
' Function Name:FolderWithVBAProjectFiles
'
' Function Purpose:
' Returns a string containing the full path name of the folder where the Forms,
' Modules, and Classes will be copied to or from.
' The folder is called VBAProjectFiles.
' It creates the folder if it doesn't exist.
'
' Called by:
' ExportModules
' ImportModules
'
' Return value:
' String containing the full path name of the VBAProjectFiles folder
'
'*******************************************************************************
'
Dim WshShell As Object
Dim FSO As Object
Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    If Not FSO.FolderExists(Path) Then
        On Error Resume Next
        MkDir Path
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(Path) Then
        FolderWithVBAProjectFiles = Path
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function ' FolderWithVBAProjectFiles

Private Sub DeleteVBAModulesAndUserForms(ByVal WkBkName As String)
'
'*******************************************************************************
' Function Name: DeleteVBAModulesAndUserForms
'
' Sub Purpose:
' Delete all the Forms, Modules, and Classes from the active workbook
' Called by:
' ImportModules
'
'*******************************************************************************
'
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent

    Set VBProj = Workbooks(WkBkName).VBProject

    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            'This is a workbook or worksheet module, we do nothing
        Else
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp

End Sub      ' DeleteVBAModulesAndUserForms

Function GetExcelFilePathAndName(ByVal Descrip As String, ByVal Criteria As String) As String
'
' Uses the Windows File Selection window
' Returns a full pathname to an Excel file
' or
' Returns an empty string ("") if no file selected
'
' Descrip is the text to appear in the file selection window header
'
' The options for Criteria are from the Drop Down box in the Save As dialog. Examples:
' Const MacroFile = "Excel Macro-Enabled Workbook(*.xlsm), *.xlsm"
' Const TextFile= "Text Files (*.txt),*.txt"
' Const AddInFile="Add-In Files (*.xla),*.xla"
' Google GetOpenFileName for more examples
'
Dim FS As Variant
    FS = Application.GetOpenFilename(FileFilter:=Criteria, MultiSelect:=False, Title:=Descrip)
    If FS <> False Then
        GetExcelFilePathAndName = FS
    Else
        GetExcelFilePathAndName = ""
    End If
    
End Function

Function GetFileName(ByVal FullPath As String) As String
Dim StrFind As String
Dim I As Integer
    Do Until Left(StrFind, 1) = "\" Or Left(StrFind, 1) = "/"
        I = I + 1
        StrFind = Right(FullPath, I)
            If I = Len(FullPath) Then ' no "\" found
                GetFileName = FullPath
                Exit Function
            End If
    Loop
    GetFileName = Right(StrFind, Len(StrFind) - 1)
End Function

Function GetFolder(ByVal Descrip As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = Descrip
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

