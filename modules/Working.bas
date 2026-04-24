Option Compare Database
Option Explicit
Public conn_err_counter As Long
Public Const strBasePathWorking As String = "F:\R i s k P o i n t\Denmark\35 M&A\M&A - Working\"
Public Const strBasePathBound As String = "F:\R i s k P o i n t\Denmark\35 M&A\M&A - Bound\"
Public Const strBasePathNonBound As String = "F:\R i s k P o i n t\Denmark\35 M&A\M&A - Non-Bound\"
Public Const strLogPath As String = "F:\R i s k P o i n t\Norway\35 M&A - Norway\Intranet\DataBase\Logs\"
Public folder_check_shall_stop As Boolean
Public moved_folder_counter As Long

Public Sub entry()
    Const proc_name As String = "working.StartFolderCheckGeneral"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    MsgBox "Hi! " & vbNewLine & vbNewLine & "I, Stella, will now look through our shared drive and backup relevant folders and files. This will take a while. Please leave me alone until I'm done. It'll look like I'm not responsive or have 'frozen', but actually I'm just busy working." _
    & "When I'm done you'll see a log which shows all the hard work I've done." & vbNewLine & vbNewLine & "Best regards, " & vbNewLine & "Stella", vbInformation, "I'll be working for a while now."
    
    open_forms.working_on_it_f "Working on it!"
    Working.folder_check_shall_stop = False
    Working.moved_folder_counter = 0
    Working.entry__working
    If folder_check_shall_stop = False Then Working.entry__bound
    If folder_check_shall_stop = False Then Working.entry__non_bound
    
    FollowHyperlink Load.system_info.system_paths.log_for_folder_moving
    DoCmd.Close acForm, "working_on_it_f"
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub entry__working()
    Const proc_name As String = "working.StartFolderCheckWorking"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    ' This module loops thorugh all files in the working folder and backs them up if not backed up already.
    ' The system is made up of two routines; this ones, which fires it, and the next one, which does the actual looping of folders.
    ' Logs the start of the procedure to a log.txt file. WriteText is a Function.
    Dim objWorking As Object
    Dim objFolder As Object
    
    Working.WriteText vbNewLine & " - - - " & vbNewLine & vbNewLine & vbNewLine & Now() & " - Folder check for working started."
    Set objWorking = CreateObject("Scripting.FileSystemObject")
        For Each objFolder In objWorking.getfolder(strBasePathWorking).SubFolders
            If folder_check_shall_stop = True Then
                WriteText Now() & " - folder_check_shall_stop = True."
                GoTo outro
            End If
            Working.folder_check objFolder
        Next objFolder
    Set objWorking = Nothing
    WriteText Now() & " - Folder check for working is completed."
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Resume outro
    End Sub
Public Sub folder_check(ByVal Path As Object)
    Const proc_name As String = "working.FolderCheckWorking"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim objFolder As Object
    Dim fso As Object
    Dim strID As String
    Dim path_actual As String
    Dim path_actual_compare As String
    Dim path_normative As String
    Dim path_normative_compare As String
    
    For Each objFolder In Path.SubFolders
        ' Need the ID to include all figures, hence string and not integer
        strID = Right(objFolder, 6)
        ' Some teams have folders which are not deals. Hopefully these don't end with four figures :)
        If IsNumeric(Right(strID, 4)) = False Then
            GoTo NextIteration
        End If
        ' Somestimes a record is deleted from Stella but the folder is not. Therefore, need to verify that dealID is in Stella.
        If DoesDealExist(Int(strID)) = 0 Then
            GoTo NextIteration
        End If
        
        path_normative = Paths.create_folder_path(Central.generate_deal_object(strID))
        If path_normative = "-1" Then
            MsgBox "Issue found: " _
                & "folder id: " & strID _
                & vbNewLine & vbNewLine & "Snip this to Christian" _
                & vbNewLine & vbNewLine & "The folder check will now stop."
                Working.folder_check_shall_stop = True
                WriteText Now() & " - folder_check_shall_stop = True"
                WriteText Now() & " - check stopped at " & strID
            GoTo outro
        End If
        
        path_normative_compare = Split(path_normative, "R i s k P o i n t")(1)
        path_actual_compare = Split(objFolder, "R i s k P o i n t")(1)
        
        If path_normative_compare <> path_actual_compare Then
            ' Logs the changes in case they need to be reversed or reviewed.
            WriteText vbNewLine & Now() & " Attempting to move from " & objFolder.Path & " to " & vbNewLine & path_normative
            ' Check whether there is a folder at the destination already. If yes, there is duplicate deal folders and human intervenation is needed.
            If Dir(path_normative, vbDirectory) = "" Then
                If Working.moved_folder_counter > 50 Then
                    Working.folder_check_shall_stop = True
                    Exit For
                End If
                If Load.is_debugging = True Then
                    Debug.Print "path_normative_compare = " & vbTab & path_normative_compare
                    Debug.Print "path_actual_compare = " & vbTab & vbTab & path_actual_compare
                    Debug.Print ""
                End If
                Set fso = CreateObject("scripting.filesystemobject")
                
                ' In case folder is in use or otherwise not accessible
                On Error Resume Next
                fso.movefolder objFolder.Path, path_normative
                On Error GoTo err_handler
                If Load.is_debugging = True Then On Error GoTo 0
    
                Working.moved_folder_counter = Working.moved_folder_counter + 1
                
                If Dir(path_normative, vbDirectory) = "" Then
                    WriteText vbNewLine & Now() & " Move failed!"
                Else
                    WriteText vbNewLine & Now() & " Move succeded!"
                End If
                fso.Close
                Set fso = Nothing
            Else
                WriteText vbNewLine & Now() & " Folder was already there. Duplicate folder warning!"
            End If
        End If

NextIteration:
    Next objFolder

outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub entry__bound()
    Const proc_name As String = "working.StartFolderCheckBound"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    ' This set of modules loop thorugh all files in the bound folder and backs them up if not backed up already.
    Dim strYear As String
    Dim i As Long
    Dim objWorking As Object
    Dim objFolder As Object
    ' The system is made up of two routines; this ones, which fires it, and the next one, which does the actual looping of folders.
    ' Logs the start of the procedure to a log.txt file. WriteText is a Function.
    
    ' Logs what happens.
    WriteText vbNewLine & " - - - " & Now() & " - Folder check for bound folders started."
    For i = Year(Date) - 2 To Year(Date)
        Set objWorking = CreateObject("Scripting.FileSystemObject")
            For Each objFolder In objWorking.getfolder(strBasePathBound & i & "/").SubFolders
                If folder_check_shall_stop = True Then
                    WriteText Now() & " - folder_check_shall_stop = True."
                    GoTo outro
                End If
                Working.folder_check objFolder
            Next objFolder
        Set objWorking = Nothing
    Next i
    WriteText Now() & " - Folder check for bound folders is completed."
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub entry__non_bound()
    Const proc_name As String = "working.StartFolderCheckNonBound"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    ' This set of modules loop thorugh all files in the bound folder and backs them up if not backed up already.
    Dim strYear As String
    Dim i As Long
    Dim y As Long
    Dim objWorking As Object
    Dim objFolder As Object
    ' The system is made up of two routines; this ones, which fires it, and the next one, which does the actual looping of folders.
    ' Logs the start of the procedure to a log.txt file. WriteText is a Function.
    
    ' Logs what happens.
    WriteText vbNewLine & " - - - " & Now() & " - Folder check for non-bound folders started."
    y = Year(Date) - 2
    If y < 2020 Then
        y = 2020
    End If
    For i = y To Year(Date)
        Set objWorking = CreateObject("Scripting.FileSystemObject")
            For Each objFolder In objWorking.getfolder(strBasePathNonBound & i & "/").SubFolders
                If folder_check_shall_stop = True Then
                    WriteText Now() & " - folder_check_shall_stop = True."
                    GoTo outro
                End If
                Working.folder_check objFolder
            Next objFolder
        Set objWorking = Nothing
    Next i
    WriteText Now() & " - Folder check for non-bound folders is completed." & vbNewLine & "_ _ _ " & vbNewLine & vbNewLine

outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Function WriteText(ByVal strText As String)
    Const proc_name As String = "working.WriteText"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim strTargetLoad As String
    strTargetLoad = Load.system_info.system_paths.log_for_folder_moving
    Open strTargetLoad For Append As #1
    Print #1, strText
    Close #1
outro:
    utilities.call_stack_remove_last_item
    Exit Function
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Function
    
Public Function DoesDealExist(ByVal varDealID As Variant)
    Const proc_name As String = "working.DoesDealExist"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    If Working.conn_err_counter > 2 Then
        Debug.Print "working.DoesDealExist failed."
        Debug.Print 1 / 0
    End If
    
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    
    str_sql = "SELECT deal_id FROM " & sources.deals_view & " WHERE deal_id = " & varDealID
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        If rs.BOF = True And rs.EOF = True Then
            DoesDealExist = 0
        Else
            DoesDealExist = 1
        End If
    rs.Close
outro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    utilities.call_stack_remove_last_item
    Exit Function
err_handler:
    If Err.Number = -2147467259 Then
        Load.check_conn_and_variables
    End If
    conn_err_counter = conn_err_counter + 1
    Central.err_handler proc_name, Err.Number, Err.Description, "", "deal_id = " & varDealID, "", True
    Resume outro
End Function