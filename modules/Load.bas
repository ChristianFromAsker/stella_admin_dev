Option Compare Database
Option Explicit

'enums must be here, as they need to be loaded immediately upon launch
Public Const stella_uw_id = 56
Public Enum menu_list
    no = 94
    yes = 93
End Enum
Public Enum deal_statuses
    nda = 1
    submission = 481
    nbi = 2
    preferred = 3
    expensed = 4
    uw = 5
    signed = 6
    closed = 436
    declined = 7
    cancelled = 8
    lost = 9
    collapsed = 485
End Enum
Public Enum navins_homes
    canada = 604
    Denmark = 127
    dubai = -1 '23 May 2025, CK: Not licensed yet.
    Finland = 130
    germany = 126
    netherlands = 509
    norway = 128
    uk_old = 125
    uk_solutions = 603
    singapore = 511
    spain = 512
    Sweden = 129
    switzerland = 510
    usa = 435
End Enum
Public Enum rp_entities
    rp_underwriting = 93
End Enum

Public Type typ_field_type
    cmd_button As String
    label As String
    text_field As String
    header As String
End Type
Public field_type As Load.typ_field_type

'These variables must load in this sequence.
Public sources As New cls_sources
Public conn As New ADODB.Connection
Public Type typ_app_continents
    eur_asia As String
    americas As String
    global As String
End Type
Public app_continents As Load.typ_app_continents

Public Type typ_form_names
    working_on_it_f As String
End Type
Public form_names As Load.typ_form_names
Public secondary_access_app As Access.Application

'general objects
Public call_stack As String
Public colors As New cls_colors
Public country_list() As Variant
Public current_uw As New cls_underwriter
Public event_id As String
Public form_backgrounds As New cls_images
Public is_debugging As Boolean
Public is_init As Boolean
Public system_info As New cls_system
Public underwriters() As Variant

Public Sub start_stella_admin()
    Load.system_info.init_system
    Load.check_conn_and_variables
    DoCmd.OpenForm "MenuMainF"
End Sub
Public Sub init_global_variables()
    is_init = True
        
    If Load.system_info.is_init = False Then Load.system_info.init_system
    
    global_vars.init
    
    Load.current_uw.init_uw
    
    With Load.form_names
        .working_on_it_f = "working_on_it_f"
    End With
    
    'variables not required for the start-up procedure
    global_vars.init_conn_dependant
    
End Sub
Public Sub check_secondary_access_app()
    Load.call_stack = Load.call_stack & vbNewLine & "load.check_secondary_access_app"
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_milestone As String
    
    'create access app for standby.
    Dim secondary_app_name As String, init_secondary_database As Boolean
    init_secondary_database = False
    str_milestone = "Before If load.secondary_access_app Is Nothing Then"
    If Load.secondary_access_app Is Nothing Then
        'app object is destroyed and must be recreated.
        init_secondary_database = True
    Else
        'even if the app object exists, the actual app might be closed due to external influence.
        'Therefore, need to check if the app is available via a try-catch solution
        secondary_app_name = ""
        On Error Resume Next
            secondary_app_name = Load.secondary_access_app.CurrentDb.Name
        On Error GoTo err_handler
        If Load.is_debugging = True Then On Error GoTo 0
        If secondary_app_name = "" Then
            init_secondary_database = True
        End If
    End If
    
    Dim str_working_on_it As String, close_working_on_it_when_done As Boolean
    close_working_on_it_when_done = False
    
    str_milestone = "Before If init_secondary_database = True Then"
    If init_secondary_database = True Then
        If CurrentProject.AllForms("working_on_it_f").IsLoaded = False Then
            str_working_on_it = "Sorry, I just need a minute."
            close_working_on_it_when_done = True
            open_forms.working_on_it_f str_working_on_it
        End If
        
        Set Load.secondary_access_app = CreateObject("Access.Application")
        With Load.secondary_access_app
            .OpenCurrentDatabase Load.system_info.system_paths.common_path & "placeholder.accdb", False
            .Visible = False
        End With
        
        If close_working_on_it_when_done = True Then open_forms.working_on_it_f__close
    End If

outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "load.check_secondary_access_app"
        .milestone = str_milestone
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub exit_stella_admin()
    If Not colors Is Nothing Then Set colors = Nothing
    If Not conn Is Nothing Then Set conn = Nothing
    If Not current_uw Is Nothing Then Set current_uw = Nothing
    If Not sources Is Nothing Then Set sources = Nothing
    
    Load.remove_images_from_access_file
    
End Sub

Public Sub init_conn()
    Const proc_name As String = "load.init_conn"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim database_name As String
    Dim encrypted_string As String
    Dim encryption_key As Long
    Dim file_path As String
    Dim ip_address As String
    Dim pw_file As Long
    Dim pwd As String
    Dim server_name As String
    Dim str_conn As String
    Dim user_name As String
    
    'default values
    database_name = Load.system_info.database_name
    encryption_key = 33
    ip_address = "mysql01-weu-prd.rpgroup.com"
    server_name = "public_avd_db"
    user_name = "stella"
    
    'get password
    pw_file = FreeFile
    
    file_path = Load.system_info.system_paths.pws & "\" & server_name & ".txt"
    Open file_path For Input As FreeFile
        encrypted_string = Input(LOF(pw_file), pw_file)
        pwd = CStr(utilities.decrypt_string(encrypted_string, Left(encryption_key, 1), Right(encryption_key, 1)))
    Close pw_file
    
    'activate connection
    If Not Load.conn Is Nothing Then Set Load.conn = Nothing
    
    str_conn = "Driver={MySQL ODBC 9.4 Unicode Driver};" _
        & "Server=" & ip_address & ";" _
        & "DATABASE=" & database_name & ";" _
        & "UID=" & user_name & ";" _
        & "PWD=" & pwd
    Set Load.conn = New ADODB.Connection
    
    With conn
        .ConnectionString = str_conn
        .Open
        .CursorLocation = adUseClient
    End With
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "init_conn failed", True
    GoTo outro
End Sub
Public Sub check_conn_and_variables()
    Const proc_name As String = "check_conn_and_variables"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    If conn Is Nothing Then Load.init_conn
    
    'connections often drops while conn.state remain 1. The .close and .open seems to fix that.
    If conn.State <> adStateClosed Then
        conn.Close
    End If
    On Error GoTo conn_fix
    conn.Open
    On Error GoTo err_handler
    If is_init = False Then Load.init_global_variables
outro:
    Exit Sub
conn_fix:
    Debug.Print Time & " conn.open failed and was reinitialised"
    Load.init_conn
    Resume Next
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: load.check_conn_and_variables" & vbNewLine _
        & "Parameters: n/a " & vbNewLine _
        & "App: Stella Admin", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub

Public Sub remove_images_from_access_file()
    'Stella adds images when loading. The images vary with seasons etc.
    'Access saves a new version of the image every time Stella is opened. Hence, the file bloats. This sub removes the bloat, as long as the database is set to 'compact on close'.
    
    Dim rs As dao.Recordset
    Dim str_sql As String
    str_sql = "SELECT * FROM MSysResources ORDER BY Id"
    Set rs = CurrentDb.OpenRecordset(str_sql)
    With rs
        .OpenRecordset
        Do Until .EOF
            If InStr(1, rs!Name, "main_menu") Then
                .Delete
            End If
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub
Public Sub init_country_list_array()
    Const proc_name As String = "load.init_country_list_array"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    str_sql = "SELECT jurisdiction_id, jurisdiction FROM " & sources.jurisdictions_view & " WHERE jurisdiction_type = 'country' ORDER BY jurisdiction"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    'rs. Open
        i = 1
        ReDim country_list(0 To CLng(rs.RecordCount), 0 To 1)
        Do Until rs.EOF
            If rs!jurisdiction <> "_All" Then
                country_list(i, 0) = rs!jurisdiction_id
                country_list(i, 1) = rs!jurisdiction
                i = i + 1
            End If
            rs.MoveNext
        Loop
        country_list(0, 0) = i - 1
    rs.Close
    
outro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub init_array_underwriters()
    Const proc_name As String = "load.init_array_underwriters"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim str_sql As String
        
    str_sql = "SELECT uw_id" _
    & ", can_change_budget_home_id" _
    & ", budget_continent, budget_continent_id" _
    & ", budget_home, budget_home_id, budget_region, budget_region_id" _
    & ", can_change_budget_home_id" _
    & ", has_admin_access_id" _
    & ", is_dev_id, is_employed_id" _
    & ", user_name, user_type_id, uw_initials, uw_name, nickname" _
    & " FROM " & Load.sources.uws_view _
    & " ORDER BY uw_initials"
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        '. Open
        ReDim underwriters(0 To CLng(rs.RecordCount), 0 To 17)
        underwriters(0, 0) = CLng(rs.RecordCount)
        i = 1
        Do Until .EOF = True
            underwriters(i, 1) = !uw_id
            underwriters(i, 2) = !uw_initials
            underwriters(i, 3) = !uw_name
            underwriters(i, 4) = !user_name
            underwriters(i, 5) = !has_admin_access_id
            underwriters(i, 6) = Nz(!budget_home, -1)
            underwriters(i, 7) = Nz(!budget_home_id, -1)
            underwriters(i, 8) = Nz(!budget_region, -1)
            underwriters(i, 9) = Nz(!budget_region_id, -1)
            underwriters(i, 10) = Nz(!nickname, !uw_name)
            underwriters(i, 11) = Nz(!budget_continent, -1)
            underwriters(i, 12) = Nz(!budget_continent_id, -1)
            underwriters(i, 13) = Nz(!can_change_budget_home_id, -1)
            underwriters(i, 14) = Nz(!is_dev_id, -1)
            underwriters(i, 15) = Nz(!is_employed_id, -1)
            underwriters(i, 16) = Nz(!user_type_id, -1)

            i = i + 1
            rs.MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub