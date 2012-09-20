Attribute VB_Name = "Buttons"
Option Explicit

'
' Author: Erik L. Eidt
' Copyright (c) 2012, All rights reserved.
' Created: 09-08-2012
'

Private Const vbext_ct_StdModule = 1
Private Const vbext_ct_ClassModule = 2
Private Const vbext_ct_MSForm = 3
Private Const vbext_ct_Document = 100

Sub VBACodeImportExport_ExportToFolder()
    Dim wkb As Workbook
    Set wkb = ActiveWorkbook
    
    Dim vbp As Variant      ' VBIDE.VBProject
    Set vbp = wkb.VBProject ' Application.Document.VBProject
    
    Dim dateTime As String
    dateTime = " on " & Format(Now(), "MM-DD-YY at HH-MM-SS")
    dateTime = Replace(dateTime, "/", "-")
    dateTime = Replace(dateTime, ":", "-")
    
    Dim exportPath As String
    
    Dim ans As String
    ans = vbNo
    If wkb.path <> "" Then
        ans = MsgBox("Exporting from vba project: " & vbp.Name & vbCr & "Use project folder for export?" & vbCr & vbCr & "Yes: use project directory" & vbCr & "(" & wkb.path & "\SourceExports\<timestamp>)" & vbCr & vbCr & "No: select another folder...", vbYesNoCancel, "Use project location for export?")
    End If
    If ans = vbCancel Then
        Exit Sub
    End If
    If ans = vbYes Then
        exportPath = CreateDirectory(wkb.path, "SourceExports")
    Else
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Choose SourceExports Directory (exports will go to timestamped folder name here)"
            .ButtonName = "Open"
            .Show
            exportPath = ""
            On Error Resume Next
            exportPath = .SelectedItems.Item(1)
            If exportPath = "" Then
                Exit Sub
            End If
        End With
    End If
    
    If exportPath = "" Then
        ans = MsgBox("Could not create: SourceExports at" & vbCr & ThisWorkbook.path, vbOKOnly, "Export")
    Else
        exportPath = CreateDirectory(exportPath, vbp.Name & dateTime)
        If exportPath = "" Then
            ans = MsgBox("Could not create: vbp.Name & dateTime" & vbCr & exportPath, vbOKOnly, "Export")
        Else
            ans = MsgBox("Exporting" & vbCr & vbCr & "from vba project: " & vbp.Name & vbCr & vbCr & "to folder:" & vbCr & """" & exportPath & """", vbOKCancel, "Export")
            
            If ans = vbOK Then
                Dim cnt As Long
                cnt = 0
                
                Dim vbi As Variant 'VBIDE.VBComponent
                For Each vbi In vbp.VBComponents
                    Dim suffix As String
                    suffix = ""
                    Select Case vbi.Type
                    Case vbext_ct_MSForm
                        suffix = ".frm"
                    Case vbext_ct_StdModule
                        suffix = ".bas"
                    'Case vbext_ct_Document
                    Case vbext_ct_ClassModule
                        suffix = ".cls"
                    End Select
                    If suffix <> "" Then
                        vbi.Export exportPath & "\" & vbi.Name & suffix
                        cnt = cnt + 1
                    End If
                Next
                MsgBox "Exported " & cnt & " files" & vbCr & vbCr & "to folder:" & vbCr & """" & exportPath & """" & vbCr & vbCr & "from vba project: " & vbp.Name, vbOKOnly, "Success"
            End If
        End If
    End If
End Sub

Sub VBACodeImportExport_ImportFromFolder()
    Dim vbp As Variant
    Set vbp = ActiveWorkbook.VBProject ' Application.Document.VBProject
    
    Dim ans As String
    ans = MsgBox("Import folder into vba project: " & vbp.Name & vbCr & "(" & vbp.fileName & ")" & vbCr & vbCr & "Select folder to import using the next dialog box.", vbOKCancel, "Import")
    If ans = vbOK Then
        Dim path As String
        'path = Application.GetOpenFilename
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Choose SourceExports Directory (exports will go to timestamped folder name here)"
            .ButtonName = "Open"
            .Show
            path = ""
            On Error Resume Next
            path = .SelectedItems.Item(1)
        End With
            
        If path <> "" And path <> "False" Then
            Dim vbc As Variant
            Set vbc = vbp.VBComponents
            
            Dim vbi As Variant
            Dim fso As Variant
            Dim ff As Variant
            Dim fl As Variant
            Dim fi As Variant
            
            Dim cnt As Long
            cnt = 0
            
            Set fso = CreateObject("Scripting.FileSystemObject")
            
            path = TrimToPath(path)
            Set ff = fso.GetFolder(path)
            Set fl = ff.Files
            
            For Each fi In fl
                Dim modName As String
                Dim suffix As String
                modName = TrimToSuffix(fi.Name, suffix)
                If suffix = ".frm" Or suffix = ".cls" Or suffix = ".bas" Then
                    Set vbi = Nothing
                    On Error Resume Next
                    Set vbi = vbc(modName)
                    On Error GoTo 0
                    If vbi Is Nothing Then
                        Set vbi = vbc.Import(path & "\" & fi.Name)
                        cnt = cnt + 1
                    Else
                        ans = MsgBox(modName & " already exists in VBA Project", vbOKCancel, "Error")
                        If ans = vbCancel Then
                            Exit For
                        End If
                    End If
                End If
            Next
            MsgBox "Imported " & cnt & " files from:" & vbCr & path & vbCr & "into vba project: " & vbp.Name, vbOKOnly, "Success"
        End If
    End If
End Sub

Function CreateDirectory(filePath As String, fileName As String) As String
    Dim d As String
    d = filePath & "\" & fileName
    If Dir(d, vbDirectory) = "" Then
        On Error GoTo E1
        MkDir d
    End If
    CreateDirectory = d
    Exit Function
E1: CreateDirectory = ""
End Function

Function TrimToPath(fname As String) As String
    Dim p As Long
    
    p = InStr(1, fname, "\")
    If p > 0 Then
        Dim q1 As Long
        q1 = p
        
        Dim q2 As Long
        Do
            q2 = InStr(q1 + 1, fname, "\")
            If q2 = 0 Then
                Exit Do
            End If
            q1 = q2
        Loop
        TrimToPath = Mid(fname, 1, q1 - 1)
    Else
        TrimToPath = ""
    End If
End Function

Function TrimToSuffix(fname As String, suffixOut As String) As String
    Dim p As Long
    
    p = InStr(1, fname, ".")
    If p > 0 Then
        Dim q1 As Long
        q1 = p
        
        Dim q2 As Long
        Do
            q2 = InStr(q1 + 1, fname, ".")
            If q2 = 0 Then
                Exit Do
            End If
            q1 = q2
        Loop
        TrimToSuffix = Mid(fname, 1, q1 - 1)
        suffixOut = Mid(fname, q1)
    Else
        TrimToSuffix = fname
        suffixOut = ""
    End If
End Function






