Attribute VB_Name = "mdl_ExportDbObjects"
Option Compare Database
Option Explicit

Public blnCurrentForm As Boolean
Const currentForm As String = "frmSourcecode Handling"
Const baseModule As String = "mdl_ExportDbObjects"

Public Sub ExportDatabaseObjects(Optional ByVal strFolder As String = "")
    On Error GoTo Err_ExportDatabaseObjects
    
    Dim db As Database
    Dim td As TableDef
    Dim d As Document
    Dim c As Container
    Dim i As Integer
    
    Dim Fs As Object
    
    Set db = CurrentDb()
    
    If strFolder = "" Then
        strFolder = Application.CurrentProject.Path & "\sourcecode\"
        Set Fs = CreateObject("Scripting.FileSystemObject")
        If Fs.FolderExists(strFolder) = False Then
            Fs.CreateFolder (strFolder)
        End If
    End If
    Kill strFolder & "*.*"
    
    For Each td In db.TableDefs
        If Left(td.Name, 4) <> "MSys" And IsTableLinked(td.Name) = False Then
            Application.ExportXML ObjectType:=acExportTable, DataSource:=td.Name, DataTarget:=(strFolder & "Table_" & td.Name & ".xml"), otherflags:=acEmbedSchema
        End If
    Next td
    
    Set c = db.Containers("Forms")
    For Each d In c.Documents
        Application.SaveAsText acForm, d.Name, strFolder & "Form_" & d.Name & ".txt"
    Next d
    
    Set c = db.Containers("Reports")
    For Each d In c.Documents
        Application.SaveAsText acReport, d.Name, strFolder & "Report_" & d.Name & ".txt"
    Next d
    
    Set c = db.Containers("Scripts")
    For Each d In c.Documents
        Application.SaveAsText acMacro, d.Name, strFolder & "Macro_" & d.Name & ".txt"
    Next d
    
    Set c = db.Containers("Modules")
    For Each d In c.Documents
        Application.SaveAsText acModule, d.Name, strFolder & "Module_" & d.Name & ".txt"
    Next d
    
    For i = 0 To db.QueryDefs.Count - 1
        Application.SaveAsText acQuery, db.QueryDefs(i).Name, strFolder & "Query_" & db.QueryDefs(i).Name & ".txt"
    Next i
    
    Set db = Nothing
    Set c = Nothing
    
    MsgBox "All database objects have been exported as a text file to " & strFolder, vbInformation

Exit_ExportDatabaseObjects:
    Exit Sub
    
Err_ExportDatabaseObjects:
    MsgBox Err.Number & " - " & Err.Description
    Resume Next
End Sub

Sub RestoreDatabaseObjectsFromFolder(Optional ByVal strFolder As String = "")
    Dim strFile As String
    Dim strSplit() As String
    Dim Fd As FileDialog
    Dim Fs As Object
    
    Set Fs = CreateObject("Scripting.FileSystemObject")
    
    If strFolder = "" Then
        strFolder = Application.CurrentProject.Path & "\sourcecode\"
    End If
 
    If Fs.FolderExists(strFolder) = False Then
        Set Fd = Application.FileDialog(msoFileDialogFolderPicker)
        With Fd
            .AllowMultiSelect = False
            .Title = "Please select folder"
            If .Show = True Then
                strFolder = .SelectedItems(1)
            Else
                Exit Sub
            End If
        End With
        Set Fd = Nothing
    End If
    
    clearDebugConsole
    
    strFile = Dir(strFolder & "*")
    Do While Len(strFile) > 0
        If InStr(strFile, ".txt") > 0 Then
            strSplit = Split(Replace(strFile, ".txt", ""), "_", 2)
        ElseIf InStr(strFile, ".xml") > 0 Then
            strSplit = Split(Replace(strFile, ".xml", ""), "_", 2)
        Else
            Debug.Print strFile & " not processed"
        End If
        Select Case strSplit(0)
            Case "Table"
                If IsTableLinked(strSplit(1)) = False Then
                    If TableExist(strSplit(1)) = True Then
                        CurrentDb.Execute "DROP TABLE " & strSplit(1), dbFailOnError
                    End If
                Application.ImportXML DataSource:=strFolder & strFile, ImportOptions:=acStructureAndData
                End If
            Case "Form"
                If blnCurrentForm = True And strSplit(1) = currentForm Then
                
                Else
                    If FormExist(strSplit(1)) = True Then
                        DoCmd.DeleteObject acForm, strSplit(1)
                    End If
                    Application.LoadFromText acForm, strSplit(1), strFolder & strFile
                End If
            Case "Report"
                If ReportExist(strSplit(1)) = True Then
                    DoCmd.DeleteObject acReport, strSplit(1)
                End If
                Application.LoadFromText acReport, strSplit(1), strFolder & strFile
            Case "Query"
                If InStr(strSplit(1), "~") > 0 Then
                    GoTo NextFile
                End If
                If QueryExist(strSplit(1)) = True Then
                    DoCmd.DeleteObject acQuery, strSplit(1)
                End If
                Application.LoadFromText acQuery, strSplit(1), strFolder & strFile
            
            Case "Macro"
                If Left(strSplit(1), 1) = "~" Then
                    GoTo NextFile
                End If
                If MacroExist(strSplit(1)) = True Then
                    DoCmd.DeleteObject acMacro, strSplit(1)
                End If
                Application.LoadFromText acMacro, strSplit(1), strFolder & strFile

            Case "Module"
                If strSplit(1) = baseModule Then
                    GoTo NextFile
                End If
                If ModuleExist(strSplit(1)) = True Then
                    DoCmd.DeleteObject acModule, strSplit(1)
                End If
                Application.LoadFromText acModule, strSplit(1), strFolder & strFile
        End Select
        Debug.Print (strSplit(1))
NextFile:
        strFile = Dir
    Loop

End Sub

Sub clearDebugConsole()
    Dim i As Integer
    For i = 0 To 200
        Debug.Print ""
    Next i
End Sub

Function FormExist(ByVal strFormname As String) As Boolean
    Dim frm As Access.AccessObject

    For Each frm In Application.CurrentProject.AllForms
        If strFormname = frm.Name Then
            FormExist = True
            Exit For
        End If
    Next
End Function

Function ModuleExist(ByVal strModulename As String) As Boolean
    Dim mdl As Access.AccessObject

    For Each mdl In Application.CurrentProject.AllMacros
        If strModulename = mdl.Name Then
            ModuleExist = True
            Exit For
        End If
    Next
End Function

Function QueryExist(ByVal strQueryname As String) As Boolean
    Dim qry

    For Each qry In CurrentDb.QueryDefs
        If strQueryname = qry.Name Then
            QueryExist = True
            Exit For
        End If
    Next
End Function

Function ReportExist(ByVal strReportname As String) As Boolean
    Dim rpt

    For Each rpt In Application.CurrentProject.AllReports
        If strReportname = rpt.Name Then
            ReportExist = True
            Exit For
        End If
    Next
End Function

Function MacroExist(ByVal strMacroname As String) As Boolean
    Dim mrc

    For Each mrc In Application.CurrentProject.AllMacros
        If strMacroname = mrc.Name Then
            MacroExist = True
            Exit For
        End If
    Next
End Function

Function TableExist(ByVal strTablename As String) As Boolean
    Dim tbldef

    For Each tbldef In CurrentDb.TableDefs
        If tbldef.Name = strTablename Then
            TableExist = True
            Exit For
        End If
    Next tbldef
End Function

Function IsTableLinked(ByVal strTablename As String) As Boolean
    ' Non-linked tables have a type of 1,
    ' tables linked using ODBC have a type of 4
    ' and all other linked tables have a type of 6
    
    IsTableLinked = Nz(DLookup("Type", "MSysObjects", "Name = '" & strTablename & "'"), 0) <> 1

End Function

