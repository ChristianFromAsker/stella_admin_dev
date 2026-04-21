Option Compare Database
Option Explicit
Private Sub empty_routine()
    Const proc_name As String = "full address of routine"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    'code
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
    ' if conn might not be available
    Debug.Print "Error in " & proc_name & " - err.number: " & Err.Number & " - err.description: " & Err.Description
End Sub

Private Sub log_activity()
    Dim proc_name As String
    Dim str_sql As String
    
    str_sql = "INSERT INTO " & Load.sources.log_activity_table & "(activity, app_name, app_continent, user_name, activity_source)" _
    & " VALUES('activity'" _
    & ", '" & Load.system_info.app_name & "'" _
    & ", '" & Load.system_info.app_continent & "'" _
    & ", '" & Environ("username") & "'" _
    & ", '" & proc_name & "')"
    Load.conn.Execute str_sql
End Sub

Private Sub test_speed()
    'at start of test sequence
    Dim timer_start As Single
    timer_start = Timer
    
    'at end of test sequence
    Debug.Print "Test took: " & Timer - timer_start
    
End Sub
Private Sub call_secondary_access_app()
    Dim output As Variant
    Dim str_milestone As String
    Dim str_folder_path As String
    Dim variable_1 As Variant
    Dim variable_2 As Variant
    
    
    Load.check_secondary_access_app
    
    'change to name of other external app as required
    Central.open_external_resource_app "scripts.accdb", False, Load.system_info.system_paths.common_path & "scripts.accdb"
    
    With Load.secondary_access_app
        'if running sub
        .Run "name of routine", variable_1, variable_2
        
        'if running function
        output = CStr(.Run("name of routine", variable_1, variable_2))
        
        str_milestone = "4"
        .CloseCurrentDatabase
        
        str_milestone = "5"
        .OpenCurrentDatabase Load.system_info.system_paths.common_path & "placeholder.accdb", False
    End With
    
outro:
    Load.secondary_access_app.Visible = False
    utilities.call_stack_remove_last_item
err_handler:
    GoTo outro
End Sub