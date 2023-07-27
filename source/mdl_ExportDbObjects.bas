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
    If Len(Dir$(strFolder & "*.*")) > 0 Then
        Kill strFolder & "*.*"
    End If
    For Each td In db.TableDefs
        If Left(td.Name, 4) <> "MSys" And IsTableLinked(td.Name, Application) = False Then
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

Sub RestoreDatabaseObjectsFromFolder(Optional ByVal strFolder As String = "", Optional ByVal strDatabase As String = "", Optional ByVal ImportData As AcImportXMLOption = acStructureAndData)
    Dim strFile As String
    Dim strSplit() As String
    Dim Fd As FileDialog
    Dim Fs As Object
    Dim db As Database
    Dim app As Access.Application
    
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
                strFolder = .SelectedItems(1) & "\"
            Else
                Exit Sub
            End If
        End With
        Set Fd = Nothing
    End If
    
    If strDatabase = "" Then
        Set app = Application
    Else
        Set app = OpenDb(strDatabase)
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
                
                If TableExist(strSplit(1), app) = True Then
                    If IsTableLinked(strSplit(1), app) = False Then
                        app.CurrentDb.Execute "DROP TABLE " & strSplit(1), dbFailOnError
                    Else
                        GoTo NextFile
                    End If
                End If
                app.ImportXML DataSource:=strFolder & strFile, ImportOptions:=ImportData
                
            Case "Form"
                If blnCurrentForm = True And strSplit(1) = currentForm Then
                
                Else
                    If FormExist(strSplit(1), app) = True Then
                        app.DoCmd.DeleteObject acForm, strSplit(1)
                    End If
                    app.LoadFromText acForm, strSplit(1), strFolder & strFile
                End If
            Case "Report"
                If ReportExist(strSplit(1), app) = True Then
                    app.DoCmd.DeleteObject acReport, strSplit(1)
                End If
                app.LoadFromText acReport, strSplit(1), strFolder & strFile
            Case "Query"
                If InStr(strSplit(1), "~") > 0 Then
                    GoTo NextFile
                End If
                If QueryExist(strSplit(1), app) = True Then
                    app.DoCmd.DeleteObject acQuery, strSplit(1)
                End If
                app.LoadFromText acQuery, strSplit(1), strFolder & strFile
            
            Case "Macro"
                If Left(strSplit(1), 1) = "~" Then
                    GoTo NextFile
                End If
                If MacroExist(strSplit(1), app) = True Then
                    app.DoCmd.DeleteObject acMacro, strSplit(1)
                End If
                app.LoadFromText acMacro, strSplit(1), strFolder & strFile

            Case "Module"
                If strSplit(1) = baseModule Then
                    GoTo NextFile
                End If
                If ModuleExist(strSplit(1), app) = True Then
                    app.DoCmd.DeleteObject acModule, strSplit(1)
                End If
                app.LoadFromText acModule, strSplit(1), strFolder & strFile
        End Select
        Debug.Print (strSplit(1))
NextFile:
        strFile = Dir
    Loop
    If strDatabase <> "" Then
        app.Quit
    End If
End Sub

Sub clearDebugConsole()
    Dim i As Integer
    For i = 0 To 200
        Debug.Print ""
    Next i
End Sub

Function FormExist(ByVal strFormname As String, app As Access.Application) As Boolean
    Dim frm As Access.AccessObject

    For Each frm In app.CurrentProject.AllForms
        If strFormname = frm.Name Then
            FormExist = True
            Exit For
        End If
    Next
End Function

Function ModuleExist(ByVal strModulename As String, app As Access.Application) As Boolean
    Dim mdl As Access.AccessObject

    For Each mdl In app.CurrentProject.AllModules
        If strModulename = mdl.Name Then
            ModuleExist = True
            Exit For
        End If
    Next
End Function

Function QueryExist(ByVal strQueryname As String, app As Access.Application) As Boolean
    Dim qry

    For Each qry In app.CurrentDb.QueryDefs
        If strQueryname = qry.Name Then
            QueryExist = True
            Exit For
        End If
    Next
End Function

Function ReportExist(ByVal strReportname As String, app As Access.Application) As Boolean
    Dim rpt

    For Each rpt In app.CurrentProject.AllReports
        If strReportname = rpt.Name Then
            ReportExist = True
            Exit For
        End If
    Next
End Function

Function MacroExist(ByVal strMacroname As String, app As Access.Application) As Boolean
    Dim mrc

    For Each mrc In app.CurrentProject.AllMacros
        If strMacroname = mrc.Name Then
            MacroExist = True
            Exit For
        End If
    Next
End Function

Function TableExist(ByVal strTablename As String, app As Access.Application) As Boolean
    Dim tbldef

    For Each tbldef In app.CurrentDb.TableDefs
        If tbldef.Name = strTablename Then
            TableExist = True
            Exit For
        End If
    Next tbldef
End Function

Function IsTableLinked(ByVal strTablename As String, app As Access.Application) As Boolean
    ' Non-linked tables have a type of 1,
    ' tables linked using ODBC have a type of 4
    ' and all other linked tables have a type of 6
    
    IsTableLinked = Nz(app.DLookup("Type", "MSysObjects", "Name = '" & strTablename & "'"), 0) <> 1

End Function

Public Function OpenDb(sDb As String) As Access.Application
    On Error GoTo Error_Handler
    Dim oAccess As Object
    Set oAccess = CreateObject("Access.Application")
    With oAccess
        .OpenCurrentDatabase sDb
        Debug.Print oAccess.CurrentDb.Name
        .Visible = True
        .UserControl = True
    End With
    
    Set OpenDb = oAccess

Error_Handler_Exit:
    On Error Resume Next
    Set oAccess = Nothing
    Exit Function

Error_Handler:
    'Occurs when the db has checks (AutoExec) UserControl to block automation, doesn't stop the connection if it is already open though.
    '-2147417848   Automation error - The object invoked has disconnected from its clients.
    If Err.Number <> 91 Then
        MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Source: OpenDb" & vbCrLf & _
               "Error Description: " & Err.Description & _
               Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
               , vbOKOnly + vbCritical, "An Error has Occurred!"
    End If
    Resume Error_Handler_Exit
End Function
