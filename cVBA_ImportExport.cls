VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cVBA_ImportExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' CLASS: cVBA_ImportExport
' DEVELOPER: Franz Hardinger
' HOMEPAGE / YOUTUBE: ""
' GITHUB: https://github.com
'
' REQUIREMENTS (WINDOWS ONLY):
' Uses Late Binding for VBIDE & WScript.Shell.
' No manual VBA References (Tools -> References) required!
'
' IMPORTANT NOTE FOR USAGE:
' In Office, the following setting must be enabled once in the Trust Center:
' File -> Options -> Trust Center -> Trust Center Settings ->
' Macro Settings -> "Trust access to the VBA project object model".
'
' DESCRIPTION:
' Automated export and hybrid import of all VBA components (.bas, .cls, .frm)
' into a local "src" folder or directly into an SVN or Git repository via a
' Windows folder shortcut ("src.lnk").
' ==============================================================================
' LICENSE: MIT License
'
' Copyright (c) 2026 Franz Hardinger
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ==============================================================================
'
' REAL-WORLD USAGE (Excel, Word & Access):
'   Dim VbaSync As New cVBA_ImportExport
'   VbaSync.Export ' Exports all modules/classes/forms to the "src" folder
'   VbaSync.Import ' Imports/updates all files from "src" into the project
' ==============================================================================


' Native VBIDE Component Type constants for Late Binding
' (No reference to "Microsoft Visual Basic for Applications Extensibility 5.3" required!)
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

Private ScrDirPath As String

Public Sub Export()
    ' Exports all VBA components with more than 2 lines of code from the active host project
    ' into the resolved "src" directory using their native file extensions (.bas, .cls, .frm).
    Dim FullName As String
    Dim I As Long
    Dim CurrentVBComponents As Object
    Dim CurrentComponent As Object
    Dim I2 As Long
    Dim FileExtension As String
    Dim ExportPath As String
    
    If Not HasVBAAccess() Then Exit Sub
    
    FullName = GetFullName()
    
    setScrDir
    If ScrDirPath = "" Then Exit Sub
    
    Debug.Print ScrDirPath
    
    ' Locate the VBProject matching the current host file path
    For I = 1 To Application.VBE.VBProjects.Count
        If Application.VBE.VBProjects.Item(I).FileName = FullName Then
            Set CurrentVBComponents = Application.VBE.VBProjects.Item(I).VBComponents
            
            ' Loop through all components in the project
            For I2 = 1 To CurrentVBComponents.Count
                Set CurrentComponent = CurrentVBComponents.Item(I2)
                
                ' Skip this exporter class itself and components with 2 or fewer lines of code
                If CurrentComponent.Name <> TypeName(Me) And CurrentComponent.CodeModule.CountOfLines > 2 Then
                    Select Case CurrentComponent.Type
                        Case vbext_ct_StdModule
                            FileExtension = ".bas"
                        Case vbext_ct_ClassModule, vbext_ct_Document
                            FileExtension = ".cls"
                        Case vbext_ct_MSForm
                            FileExtension = ".frm"
                        Case Else
                            FileExtension = ""
                    End Select
                    
                    If FileExtension <> "" Then
                        ExportPath = ScrDirPath & "\" & CurrentComponent.Name & FileExtension
                        CurrentComponent.Export ExportPath
                    End If
                End If
            Next I2
            
            Exit Sub
        End If
    Next I
End Sub

Public Sub Import()
    ' Imports all .bas, .cls, and .frm files from the "src" directory into the active host project.
    ' Hybrid approach:
    ' - Existing Standard (.bas) and Document modules get their code replaced in-place.
    ' - Class (.cls) and UserForm (.frm) modules are removed and re-imported via .Import
    '   to preserve hidden VB attributes (such as VB_UserMemId = 0) and binary .frx layouts.
    Dim FullName As String
    Dim I As Long
    Dim I2 As Long
    Dim CurrentVBComponents As Object
    Dim CurrentComponent As Object
    Dim FileName As String
    Dim FileExtension As String
    Dim ComponentName As String
    Dim FilePath As String
    Dim FilesToImport() As String
    Dim FileCount As Long
    Dim ComponentExists As Boolean
    
    If Not HasVBAAccess() Then Exit Sub
    
    FullName = GetFullName()
    
    setScrDir
    If ScrDirPath = "" Then Exit Sub
    
    ' Locate the VBProject matching the current host file path
    For I = 1 To Application.VBE.VBProjects.Count
        If Application.VBE.VBProjects.Item(I).FileName = FullName Then
            Set CurrentVBComponents = Application.VBE.VBProjects.Item(I).VBComponents
            
            ' Pass 1: Collect valid files, update .bas/Document modules in-place,
            ' and remove existing .cls/.frm modules first
            FileName = Dir(ScrDirPath & "\*.*")
            Do While FileName <> ""
                If InStrRev(FileName, ".") > 0 Then
                    FileExtension = LCase(Mid(FileName, InStrRev(FileName, ".")))
                    ComponentName = Left(FileName, InStrRev(FileName, ".") - 1)
                    FilePath = ScrDirPath & "\" & FileName
                    
                    ' Process only valid VBA source files and skip this importer class itself
                    If (FileExtension = ".bas" Or FileExtension = ".cls" Or FileExtension = ".frm") _
                       And ComponentName <> TypeName(Me) Then
                        
                        ComponentExists = False
                        
                        ' Check if the component already exists in the project
                        For I2 = 1 To CurrentVBComponents.Count
                            Set CurrentComponent = CurrentVBComponents.Item(I2)
                            If LCase(CurrentComponent.Name) = LCase(ComponentName) Then
                                ComponentExists = True
                                Exit For
                            End If
                        Next I2
                        
                        If ComponentExists Then
                            If CurrentComponent.Type = vbext_ct_StdModule Or CurrentComponent.Type = vbext_ct_Document Then
                                ' Replace code lines directly without removing the module
                                ReplaceModuleCode CurrentComponent, FilePath
                            Else
                                ' Remove Class or UserForm so it can be cleanly re-imported in Pass 2
                                If Application.Name = "Microsoft Access" Then
                                    CallByName Application.DoCmd, "Rename", VbMethod, CurrentComponent.Name & "OLD", 5, CurrentComponent.Name
                                End If
                                CurrentVBComponents.Remove CurrentComponent
                                FileCount = FileCount + 1
                                ReDim Preserve FilesToImport(1 To FileCount)
                                FilesToImport(FileCount) = FilePath
                            End If
                        Else
                            ' Brand new component -> queue for .Import in Pass 2
                            FileCount = FileCount + 1
                            ReDim Preserve FilesToImport(1 To FileCount)
                            FilesToImport(FileCount) = FilePath
                        End If
                    End If
                End If
                
                FileName = Dir()
            Loop
            
            ' Allow VBE to finalize all pending removals before re-importing
            If FileCount > 0 Then
                DoEvents
                
                ' Pass 2: Import new modules and re-import removed classes/forms
                For I2 = 1 To FileCount
                    CurrentVBComponents.Import FilesToImport(I2)
                Next I2
            End If
            
            Exit Sub
        End If
    Next I
End Sub

Private Sub ReplaceModuleCode(ByVal TargetComponent As Object, ByVal FilePath As String)
    ' Reads a .bas or .cls file, strips the VB header/attributes, and replaces the target code lines.
    Dim FileNum As Integer
    Dim FileContent As String
    Dim Lines() As String
    Dim I As Long
    Dim CleanCode As String
    Dim HeaderEnded As Boolean
    
    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    FileContent = Input$(LOF(FileNum), FileNum)
    Close #FileNum
    
    Lines = Split(FileContent, vbCrLf)
    
    For I = LBound(Lines) To UBound(Lines)
        If HeaderEnded Then
            CleanCode = CleanCode & Lines(I) & vbCrLf
        ElseIf Left(Trim(Lines(I)), 10) <> "Attribute " And Left(Trim(Lines(I)), 8) <> "VERSION " _
               And Trim(Lines(I)) <> "BEGIN" And Trim(Lines(I)) <> "END" _
               And Left(Trim(Lines(I)), 8) <> "MultiUse" Then
            ' First real code line reached after the VBA file header
            HeaderEnded = True
            CleanCode = CleanCode & Lines(I) & vbCrLf
        End If
    Next I
    
    With TargetComponent.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(CleanCode) > 0 Then .AddFromString CleanCode
    End With
End Sub

Private Function GetFullName() As String
    Dim ApplicationObject As Object
    
    Select Case Application.Name
        Case "Microsoft Excel"
            GetFullName = CallByName(Application, "ThisWorkbook", VbGet).FullName
        Case "Microsoft Word"
            GetFullName = CallByName(Application, "ActiveDocument", VbGet).FullName
        Case "Microsoft Access"
            Set ApplicationObject = CallByName(Application, "CurrentDb", VbMethod)
            GetFullName = ApplicationObject.Name
    End Select
End Function

Private Function HasVBAAccess() As Boolean
    ' Checks if "Trust access to the VBA project object model" is enabled in the host application
    Dim TestCount As Long
    On Error GoTo ErrHandler
    TestCount = Application.VBE.VBProjects.Count
    HasVBAAccess = True
    Exit Function
ErrHandler:
    MsgBox "Trust access to the VBA project object model is disabled!" & vbCrLf & vbCrLf & _
           "Please enable it under: File -> Options -> Trust Center -> Trust Center Settings -> " & _
           "Macro Settings -> 'Trust access to the VBA project object model'.", _
           vbExclamation, "VBA Import/Export"
    HasVBAAccess = False
End Function

Public Sub setScrDir()
    ' Resolves and sets the target "src" directory path for VBA code export.
    ' Supports local drives, OneDrive-synced files, and Windows folder shortcuts (src.lnk).
    ' Creates a local "src" directory if neither a valid folder nor a shortcut exists.
    Dim fileFullName As String
    Dim baseDir As String
    Dim targetDir As String
    Dim shortcutFile As String
    Dim wshShell As Object
    Dim shortcutObj As Object
    Dim resolvedPath As String
    Dim oneDriveRoot As String
    
    ' Retrieve the full path of the active host document/database
    fileFullName = GetFullName()
    
    ' Convert OneDrive/SharePoint HTTPS URLs to their local sync folder path
    If Left(LCase(fileFullName), 4) = "http" Then
        oneDriveRoot = Environ("OneDrive")
        If oneDriveRoot = "" Then oneDriveRoot = Environ("OneDriveCommercial")
        
        If oneDriveRoot <> "" Then
            fileFullName = Replace(fileFullName, "/", "\")
            ' Strip the "https:\\d.docs.live.net\<UserID>" prefix to append the relative path
            fileFullName = oneDriveRoot & Mid(fileFullName, InStr(10, fileFullName, "\"))
        End If
    End If
    
    ' Abort if the file has not been saved locally yet (no directory separator found)
    If InStrRev(fileFullName, "\") = 0 Then Exit Sub
    
    ' Extract the parent directory path and define expected "src" targets
    baseDir = Left(fileFullName, InStrRev(fileFullName, "\"))
    targetDir = baseDir & "src"
    shortcutFile = baseDir & "src.lnk"
    
    ' 1. Use the local "src" directory if it already exists
    If Dir(targetDir & "\", vbDirectory) <> "" Then
        ScrDirPath = targetDir
        
    ' 2. Resolve the target path if a "src.lnk" shortcut is present
    ElseIf Dir(shortcutFile) <> "" Then
        Set wshShell = CreateObject("WScript.Shell")
        Set shortcutObj = wshShell.CreateShortcut(shortcutFile)
        resolvedPath = shortcutObj.TargetPath
        
        ' Verify that the shortcut target exists; otherwise fall back to a local "src" folder
        If resolvedPath <> "" And Dir(resolvedPath & "\", vbDirectory) <> "" Then
            ScrDirPath = resolvedPath
        Else
            MkDir targetDir
            ScrDirPath = targetDir
        End If
        
    ' 3. Create a new local "src" directory if neither folder nor shortcut exists
    Else
        MkDir targetDir
        ScrDirPath = targetDir
    End If
End Sub
