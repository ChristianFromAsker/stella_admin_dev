Option Compare Database
Option Explicit

Private Const EXPORT_ROOT_FOLDER As String = "D:\riskpoint\stella\"

Public Sub entry()
    Const proc_name As String = "export_vba.entry"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0

    Dim dbPath As String
    Dim outRoot As String
    
    open_forms.working_on_it_f "Exporting code.", "Hold on please.", 5000
    dbPath = CurrentDb.name

    outRoot = BuildExportRoot(EXPORT_ROOT_FOLDER, dbPath)

    entry__ensure_folder outRoot
    entry__ensure_folder outRoot & "\modules"
    entry__ensure_folder outRoot & "\classes"
    entry__ensure_folder outRoot & "\forms"
    entry__ensure_folder outRoot & "\forms_full"
    entry__ensure_folder outRoot & "\other"
    entry__ensure_folder outRoot & "\logs"

    ' Export code-bearing objects (modules/classes/forms/reports/macros)
    export_vba.entry__export_all outRoot, True

    ' Write a small manifest (helps auditing & AI context)
    WriteManifest outRoot & "\logs", dbPath

outro:
    utilities.call_stack_remove_last_item True
    Exit Sub
err_handler:
    MsgBox "Export failed " & proc_name & ": " & Err.Number & " - " & Err.Description, vbCritical
    Resume outro
End Sub
Public Sub import_resource()
    
    Dim dct_resource As Scripting.Dictionary
    Dim col_resources As Collection
    
    Set col_resources = New Collection
    
    Set dct_resource = New Scripting.Dictionary
    dct_resource.Add "resource_name", "utilities"
    dct_resource.Add "resource_path", "D:\riskpoint\stella\stella_uw_dev\modules\" & dct_resource("resource_name")
    dct_resource.Add "resource_type", acModule
    dct_resource.Add "resource_ending", "bas"
    col_resources.Add dct_resource
    
    Set dct_resource = New Scripting.Dictionary
    dct_resource.Add "resource_name", "cls_field"
    dct_resource.Add "resource_path", "D:\riskpoint\stella\stella_uw_dev\classes\" & dct_resource("resource_name")
    dct_resource.Add "resource_type", acModule
    dct_resource.Add "resource_ending", "cls"
    col_resources.Add dct_resource
    
    Set dct_resource = New Scripting.Dictionary
    dct_resource.Add "resource_name", "export_vba"
    dct_resource.Add "resource_path", "D:\riskpoint\stella\stella_uw_dev\modules\" & dct_resource("resource_name")
    dct_resource.Add "resource_type", acModule
    dct_resource.Add "resource_ending", "bas"
    col_resources.Add dct_resource
    
    For Each dct_resource In col_resources
        Application.LoadFromText dct_resource("resource_type"), dct_resource("resource_name"), dct_resource("resource_path") & "." & dct_resource("resource_ending")
    Next dct_resource

End Sub
Public Sub entry__export_all(ByVal outRoot As String, ByVal full_export As Boolean)
    Const proc_name As String = "export_vba.entry"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim c As AccessObject
    Dim vbComp As Object
    Dim ext As String
    Dim outPath As String

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' vbext_ct_StdModule
                ext = "bas"
                outPath = outRoot & "\modules\" & vbComp.name & "." & ext
                Application.SaveAsText acModule, vbComp.name, outPath

            Case 2 ' vbext_ct_ClassModule
                ext = "cls"
                outPath = outRoot & "\classes\" & vbComp.name & "." & ext
                Application.SaveAsText acModule, vbComp.name, outPath
            Case 100
                ext = "txt"
                outPath = outRoot & "\forms\" & vbComp.name & "." & ext
                vbComp.Export outPath
                If full_export = True Then
                    outPath = outRoot & "\forms_full\" & Split(vbComp.name, "orm_")(1) & "." & ext
                    Application.SaveAsText acForm, Split(vbComp.name, "orm_")(1), outPath
                End If
            Case Else
        End Select
    Next vbComp
outro:
    utilities.call_stack_remove_last_item True
    Exit Sub
err_handler:
    MsgBox "Export failed at " & proc_name & ": " & Err.Number & " - " & Err.Description, vbCritical
    Resume outro
End Sub
Private Function BuildExportRoot(ByVal root As String, ByVal dbPath As String) As String
    Dim dbName As String
    dbName = Mid$(dbPath, InStrRev(dbPath, "\") + 1)
    dbName = Left$(dbName, InStrRev(dbName, ".") - 1)
    BuildExportRoot = TrimTrailingSlash(root) & "\" & SafeFileName(dbName)
End Function

Private Sub entry__ensure_folder(ByVal folderPath As String)
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
End Sub

Private Function TrimTrailingSlash(ByVal p As String) As String
    If Right$(p, 1) = "\" Then
        TrimTrailingSlash = Left$(p, Len(p) - 1)
    Else
        TrimTrailingSlash = p
    End If
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    SafeFileName = s
    For Each ch In badChars
        SafeFileName = Replace(SafeFileName, CStr(ch), "_")
    Next ch
End Function

Private Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Output As #ff
    Print #ff, content
    Close #ff
End Sub

Private Sub WriteManifest(ByVal outRoot As String, ByVal dbPath As String)
    Dim s As String
    s = "ExportedFrom=" & dbPath & vbCrLf & _
        "ExportedAt=" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
        "AccessVersion=" & Application.Version & vbCrLf
    WriteTextFile outRoot & "\EXPORT_MANIFEST.txt", s
End Sub

Private Function IsSystemObject(ByVal objectName As String) As Boolean
    ' Conservative filter: skip hidden/system-ish naming
    IsSystemObject = (Left$(objectName, 4) = "MSys") Or (Left$(objectName, 1) = "~")
End Function

Private Function IsSystemTable(ByVal tableName As String) As Boolean
    IsSystemTable = (Left$(tableName, 4) = "MSys")
End Function

Private Function IsSystemQueryDef(ByVal qName As String) As Boolean
    ' Skip system/internal/temp queries
    IsSystemQueryDef = (Left$(qName, 4) = "~sq_") Or (Left$(qName, 4) = "MSys") Or (Left$(qName, 1) = "~")
End Function

Private Function FieldTypeName(ByVal daoType As Integer) As String
    ' Minimal mapper for readability
    Select Case daoType
        Case dbText: FieldTypeName = "Text"
        Case dbMemo: FieldTypeName = "LongText"
        Case dbLong: FieldTypeName = "Long"
        Case dbInteger: FieldTypeName = "Integer"
        Case dbDouble: FieldTypeName = "Double"
        Case dbSingle: FieldTypeName = "Single"
        Case dbCurrency: FieldTypeName = "Currency"
        Case dbDate: FieldTypeName = "Date"
        Case dbBoolean: FieldTypeName = "Boolean"
        Case dbGUID: FieldTypeName = "GUID"
        Case Else: FieldTypeName = "DAOType(" & daoType & ")"
    End Select
End Function