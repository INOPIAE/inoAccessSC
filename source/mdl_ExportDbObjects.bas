Attribute VB_Name = "mdl_ExportDbObjects"
Option Compare Database
Option Explicit

Public blnCurrentForm As Boolean
Const currentForm As String = "frmSourcecode Handling"
Const baseModule As String = "mdl_ExportDbObjects"

Public Sub ExportDatabaseObjects(Optional ByVal strFolder As String = "", Optional ByVal blnHash As Boolean = False)
    On Error GoTo Err_ExportDatabaseObjects
    
    Dim db As Database
    Dim td As TableDef
    Dim d As Document
    Dim c As Container
    Dim I As Integer
    Dim appVersion As String
    
    Dim fs As Object
    
    Set db = CurrentDb()
    
    If blnHash = True Then
        CreateHashTable
        appVersion = getAppVersion
    Else
        appVersion = ""
    End If
    
    If strFolder = "" Then
        strFolder = Application.CurrentProject.Path & "\sourcecode\"
        Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FolderExists(strFolder) = False Then
            fs.CreateFolder (strFolder)
        End If
    End If
    If Len(Dir$(strFolder & "*.*")) > 0 Then
        Kill strFolder & "*.*"
    End If
    For Each td In db.TableDefs
        If Left(td.Name, 4) <> "MSys" And IsTableLinked(td.Name, Application) = False Then
            Application.ExportXML ObjectType:=acExportTable, DataSource:=td.Name, DataTarget:=(strFolder & "Table_" & td.Name & ".xml"), otherflags:=acEmbedSchema
            WriteAppVersion td.Name, strFolder & "Table_" & td.Name & ".xml", appVersion
        End If
    Next td
    
    Set c = db.Containers("Forms")
    For Each d In c.Documents
        Application.SaveAsText acForm, d.Name, strFolder & "Form_" & d.Name & ".txt"
        WriteAppVersion d.Name, strFolder & "Form_" & d.Name & ".txt", appVersion
    Next d
    
    Set c = db.Containers("Reports")
    For Each d In c.Documents
        Application.SaveAsText acReport, d.Name, strFolder & "Report_" & d.Name & ".txt"
        WriteAppVersion d.Name, strFolder & "Report_" & d.Name & ".txt", appVersion
    Next d
    
    Set c = db.Containers("Scripts")
    For Each d In c.Documents
        Application.SaveAsText acMacro, d.Name, strFolder & "Macro_" & d.Name & ".txt"
        WriteAppVersion d.Name, strFolder & "Macro_" & d.Name & ".txt", appVersion
    Next d
    
    Set c = db.Containers("Modules")
    For Each d In c.Documents
        Application.SaveAsText acModule, d.Name, strFolder & "Module_" & d.Name & ".txt"
        WriteAppVersion d.Name, strFolder & "Module_" & d.Name & ".txt", appVersion
    Next d
    
    For I = 0 To db.QueryDefs.Count - 1
        Application.SaveAsText acQuery, db.QueryDefs(I).Name, strFolder & "Query_" & db.QueryDefs(I).Name & ".txt"
        WriteAppVersion db.QueryDefs(I).Name, strFolder & "Query_" & db.QueryDefs(I).Name & ".txt", appVersion
    Next I
    
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
    Dim fs As Object
    Dim db As Database
    Dim app As Access.Application
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If strFolder = "" Then
        strFolder = Application.CurrentProject.Path & "\sourcecode"
    End If
 
    If fs.FolderExists(strFolder) = False Then
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
    
    strFolder = strFolder & "\"
    If strDatabase = "" Then
        Set app = Application
    Else
        Set app = OpenDb(strDatabase)
    End If
    
    clearDebugConsole

    strFile = Dir(strFolder & "*")
    Do While Len(strFile) > 0
        If strFile = "version.txt" Then
            UpdateDBVersion strFolder & strFile, app
            GoTo NextFile
        ElseIf InStr(strFile, ".txt") > 0 Then
            strSplit = Split(Replace(strFile, ".txt", ""), "_", 2)
        ElseIf InStr(strFile, ".xml") > 0 Then
            strSplit = Split(Replace(strFile, ".xml", ""), "_", 2)
        Else
            Debug.Print strFile & " not processed"
            GoTo NextFile
        End If
        ImportFile strFolder, strSplit(1), app
        Debug.Print (strSplit(1))
NextFile:
        strFile = Dir
    Loop
    If strDatabase <> "" Then
        app.Quit
    End If
End Sub

Sub clearDebugConsole()
    Dim I As Integer
    For I = 0 To 200
        Debug.Print ""
    Next I
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


Sub CreateHashTable()
    Dim app As Access.Application
    Dim strTable As String
    strTable = "tbl_DevInfo"
    Set app = Application
    If TableExist(strTable, app) = False Then
        Dim db As DAO.Database
        Dim fld As DAO.Field
        Dim prp As DAO.Property
        Dim Tbl As DAO.TableDef
        Dim Ind As DAO.Index
        
        Set db = app.CurrentDb
        Set Tbl = db.CreateTableDef(strTable)
            
        With Tbl
            Set fld = .CreateField(strTable & "ID", dbLong)
            fld.Attributes = dbAutoIncrField + dbFixedField
            .Fields.Append fld
            
        End With
        
        'Primary key index.
        Set Ind = Tbl.CreateIndex("PrimaryKey")
        With Ind
            .Fields.Append .CreateField(strTable & "ID")
            .Unique = True
            .Primary = True
        End With
        Tbl.Indexes.Append Ind
        
        Set fld = Tbl.CreateField("ObjectName", dbText, 255)
        Tbl.Fields.Append fld
    
        Set fld = Tbl.CreateField("AppVersion", dbText, 10)
        Tbl.Fields.Append fld
        
        Set fld = Tbl.CreateField("FileHash", dbText, 128)
        Tbl.Fields.Append fld
                
        Set fld = Tbl.CreateField("created", dbDate)
        fld.DefaultValue = "=Now()"
        Tbl.Fields.Append fld
        
        'Save the table.
        db.TableDefs.Append Tbl
    End If
End Sub

Function getAppVersion() As String
    Dim app As Access.Application
    Dim strTable As String
    strTable = "tDBInfo"
    Set app = Application
    If TableExist(strTable, app) = False Then
        getAppVersion = "0.1"
    Else
        Dim strSql As String
        Dim rsDB As DAO.Recordset
        strSql = "SELECT tDBInfo.* FROM tDBInfo "
        Set rsDB = app.CurrentDb.OpenRecordset(strSql, dbOpenDynaset)
        With rsDB
            getAppVersion = .Fields(0).Value
        End With
    End If
End Function

Sub WriteAppVersion(strObject As String, strFilename As String, strAppVersion As String)
    Dim strSql As String
    Dim HashString As String
    If InStr(strFilename, "~") > 0 Or strAppVersion = "" Then
        Exit Sub
    End If
    HashString = FileToSHA512(strFilename)
    If IsNull(DLookup("tbl_DevInfoID", "tbl_DevInfo", "AppVersion = '" & strAppVersion & "' AND ObjectName = '" & strObject & "'")) Then
        strSql = "INSERT INTO tbl_DevInfo (ObjectName, AppVersion, FileHash) VALUES ('" & strObject & "', '" & strAppVersion & "', '" & HashString & "')"
        CurrentDb.Execute strSql
    Else
        Dim id As Long
        id = DLookup("tbl_DevInfoID", "tbl_DevInfo", "AppVersion = '" & strAppVersion & "' AND ObjectName = '" & strObject & "'")
        strSql = "UPDATE tbl_DevInfo SET FileHash = '" & HashString & "', created = NOW() WHERE tbl_DevInfoID = " & id
        CurrentDb.Execute strSql
    End If
    
End Sub

Public Function FileToSHA512(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA2-512 hash
    'Set a reference to mscorlib 4.0 64-bit
    'Make sure that Net Framework 3.5 (includes .Net 2 and .Net 3) is installed and enabled
    'and not only the Net Framework 4.8 Advanced Services
    
    Dim enc, bytes, outstr As String, pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.SHA512Managed")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = True Then
       FileToSHA512 = ConvToBase64String(bytes)
    Else
       FileToSHA512 = ConvToHexString(bytes)
    End If

    Set enc = Nothing

End Function

Private Function GetFileBytes(ByVal sPath As String) As Byte()
    'makes byte array from file
    'Set a reference to mscorlib 4.0 64-bit
    'Make sure that Net Framework 3.5 (includes .Net 2 and .Net 3) is installed and enabled
    'and not only the Net Framework 4.8 Advanced Services
    
    Dim lngFileNum As Long, bytRtnVal() As Byte, bTest
    
    lngFileNum = FreeFile

    If LenB(Dir(sPath)) Then ''// Does file exist?
        
        Open sPath For Binary Access Read As lngFileNum
        
        'a zero length file content will give error 9 here
        
        ReDim bytRtnVal(0 To LOF(lngFileNum) - 1&) As Byte
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    Else
        Err.Raise 53 'File not found
    End If
    
    GetFileBytes = bytRtnVal
    
    Erase bytRtnVal

End Function

Function ConvToBase64String(vIn As Variant) As Variant
    'used to produce a base-64 output
    'Set a reference to mscorlib 4.0 64-bit
    'Make sure that Net Framework 3.5 (includes .Net 2 and .Net 3) is installed and enabled
    'and not only the Net Framework 4.8 Advanced Services
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Function ConvToHexString(vIn As Variant) As Variant
     'used to produce a hex output
    'Set a reference to mscorlib 4.0 64-bit
    'Make sure that Net Framework 3.5 (includes .Net 2 and .Net 3) is installed and enabled
    'and not only the Net Framework 4.8 Advanced Services
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Sub RestoreSingleDatabaseObjects(ByVal strFile As String, Optional ByVal strDatabase As String = "", Optional ByVal ImportData As AcImportXMLOption = acStructureAndData)
    Dim strFolder As String
    Dim f
    Dim Fd As FileDialog
    Dim fs As Object
    Dim db As Database
    Dim app As Access.Application
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    

 
    If fs.FileExists(strFile) = False Then
        Set Fd = Application.FileDialog(msoFileDialogOpen)
        With Fd
            .AllowMultiSelect = False
            .Title = "Please select a file"
            If .Show = True Then
                strFile = .SelectedItems(1)
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
    
    Set f = fs.GetFile(strFile)
    strFolder = f.ParentFolder & "\"
    strFile = f.Name
    
    ImportFile strFolder, strFile, app
    
    If strDatabase <> "" Then
        app.Quit
    End If
End Sub

Public Function ImportFile(ByVal strFolder As String, ByVal strFile As String, app As Access.Application, Optional ByVal ImportData As AcImportXMLOption = acStructureAndData) As Boolean
    Dim strSplit() As String
    
    ImportFile = False
    
    If InStr(strFile, ".txt") > 0 Then
        strSplit = Split(Replace(strFile, ".txt", ""), "_", 2)
    ElseIf InStr(strFile, ".xml") > 0 Then
        strSplit = Split(Replace(strFile, ".xml", ""), "_", 2)
    Else
        Debug.Print strFile & " not processed"
        Exit Function
    End If
    Select Case strSplit(0)
        Case "Table"
            
            If TableExist(strSplit(1), app) = True Then
                If IsTableLinked(strSplit(1), app) = False Then
                    app.CurrentDb.Execute "DROP TABLE " & strSplit(1), dbFailOnError
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
                Exit Function
            End If
            If QueryExist(strSplit(1), app) = True Then
                app.DoCmd.DeleteObject acQuery, strSplit(1)
            End If
            app.LoadFromText acQuery, strSplit(1), strFolder & strFile
        
        Case "Macro"
            If Left(strSplit(1), 1) = "~" Then
                Exit Function
            End If
            If MacroExist(strSplit(1), app) = True Then
                app.DoCmd.DeleteObject acMacro, strSplit(1)
            End If
            app.LoadFromText acMacro, strSplit(1), strFolder & strFile

        Case "Module"
            If strSplit(1) = baseModule Then
                Exit Function
            End If
            If ModuleExist(strSplit(1), app) = True Then
                app.DoCmd.DeleteObject acModule, strSplit(1)
            End If
            app.LoadFromText acModule, strSplit(1), strFolder & strFile
    End Select
    Debug.Print (strSplit(1))
    ImportFile = True

End Function

Public Function UpdateDBVersion(ByVal strFile As String, app As Access.Application)
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")

    Dim verFile As Object
    Dim myFilePath As String
    Dim strInput As String
    Dim appVersion() As String
    Dim strSql As String

    Const ForReading = 1
   
    Set verFile = fs.OpenTextFile(strFile, ForReading)
    strInput = verFile.ReadAll
    verFile.Close
    
    appVersion = Split(strInput, ":")
    
        
    strSql = "Update tDBInfo SET AnVer = '" & appVersion(1) & "', Zustand = FALSE"
    app.CurrentDb.Execute strSql

End Function


