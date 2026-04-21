Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public col_all_fields As New Collection

Public cmd_show_devs As New cls_field
Public cmd_filter_closed_issues As New cls_field
Public connection_count As New cls_field
Public connection_max As New cls_field
Public is_init As Boolean
Public filter_app_continent As New cls_field
Public filter_app_name As New cls_field
Public filter_routine_name As New cls_field
Public form_name As String

Public Sub init_collections()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_form_system_dashboard.init_collections"
    Set col_all_fields = Nothing
    With col_all_fields
        .Add Me.connection_count
        .Add Me.connection_max
        .Add Me.cmd_show_devs
        .Add Me.cmd_filter_closed_issues
        .Add Me.filter_app_name
        .Add Me.filter_app_continent
        .Add Me.filter_routine_name
    End With
End Sub

Public Sub init()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_form_system_dashboard.init"
    form_name = "system_dashboard_f"
    Me.init_collections
    With Me.connection_count
        .field_name = "connection_count"
    End With
    
    With Me.connection_max
        .field_name = "connection_max"
    End With
    
    With Me.cmd_show_devs
        .field_name = "cmd_dev_filter"
        .field_value = False
    End With
    
    With Me.cmd_filter_closed_issues
        .field_name = "cmd_filter_closed_issues"
        .field_value = False
    End With
    
    With Me.filter_routine_name
        .field_name = "header_routine_name"
        .field_value = "_all"
    End With
    With Me.filter_app_continent
        .field_name = "header_app_continent"
        .field_value = "_all"
    End With
    With Me.filter_app_name
        .field_name = "header_app_name"
        .field_value = "_all"
    End With
    Me.is_init = True
End Sub

Public Sub refresh()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_form_system_dashboard.refresh"
    If CurrentProject.AllForms(Me.form_name).IsLoaded = False Then
        GoTo outro
    End If
    
    With Forms(Me.form_name)
        .Controls(Me.connection_count.field_name).Caption = "Current database connections: " & Me.get_connection_count
        .Controls(Me.connection_max.field_name).Caption = "Max database connections: " & Me.get_connection_max
    End With
    
    fix_rs.system_dashboard_f
    
outro:
    Exit Sub
    
End Sub
Public Function get_connection_count()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_form_system_dashboard.get_connection_count"
    Dim str_sql As String, rs As ADODB.Recordset
    str_sql = "SHOW status WHERE `variable_name` = 'Threads_connected'"
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        get_connection_count = rs!Value
    rs.Close
    Set rs = Nothing
End Function
Public Function get_connection_max()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_form_system_dashboard.get_connection_max"
    Dim str_sql As String, rs As ADODB.Recordset
    str_sql = "Show Variables Like 'max_connections'"
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        get_connection_max = rs!Value
    rs.Close
    Set rs = Nothing
End Function
Public Sub paint()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_form_system_dashboard.paint"
    Dim form_field As cls_field
    For Each form_field In Me.col_all_fields
        With form_field
            If .field_caption <> "-1" Then
                Forms(Me.form_name).Controls(.field_name).Caption = .field_caption
            End If
        End With
    Next form_field
End Sub
Public Sub update_cm(ByVal input_app_continent As String)
    Const proc_name As String = "cls_form_system_dashboard.update_cm"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.system_info.system_paths.init
    
    Const yes As Integer = 93
    Const budget_home_us As Long = 93
    
    Dim fso As Object
    Dim source_path As String
    Dim target_path As String
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    
    Set fso = CreateObject("scripting.filesystemobject")
    
    source_path = Load.system_info.system_paths.stella_path & "stable_builds\cm_uw.accdb"
    
    'save new version to back-up folder
    target_path = Load.system_info.system_paths.old_stella_versions & Year(Date) & Month(Date) & Day(Date) & "-" & Hour(Time) & Minute(Time) & Second(Time) & " cm_uw.accdb"
    Call fso.CopyFile(source_path, target_path, True)
    
    'save new version to stable_folder in Americas
    target_path = Load.system_info.system_paths.stable_builds_us & "cm_uw.accdb"
    Call fso.CopyFile(source_path, target_path, True)
    
    'copy stable build to production area
    str_sql = "SELECT uw_id, uw_initials, budget_home_id, is_dev_id FROM " & sources.uws_view & " WHERE is_employed_id = " & menu_list.yes
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        Do Until rs.EOF = True
            
            If input_app_continent = Load.system_info.continents.americas Then
                target_path = Load.system_info.system_paths.individual_path_us & "cm - " & !uw_initials & " " & rs!uw_id & ".accdb"
                Call fso.CopyFile(source_path, target_path, True)
            End If
            
            If input_app_continent = Load.system_info.continents.eurasia Then
                target_path = Load.system_info.system_paths.individual_path_eur & "cm - " & !uw_initials & " " & rs!uw_id & ".accdb"
                Call fso.CopyFile(source_path, target_path, True)
            End If
            
            .MoveNext
            
        Loop
        .Close
    End With
    Set rs = Nothing
    Set fso = Nothing
    
outro:
    open_forms.working_on_it_f__close
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "cls_form_system_dashboard.update_cm"
        .milestone = "str_sql = " & str_sql
        .params = "input_app_continent = " & input_app_continent
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With

End Sub
Public Sub update_stella_uw(ByVal input_app_continent As String)
    Const proc_name As String = "cls_form_system_dashboard.update_stella_uw"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.system_info.system_paths.init
    
    Dim fso As Object
    Dim rs As ADODB.Recordset
    Dim source_path As String
    Dim source_path_placeholder_eur As String
    Dim source_path_placeholder_usa As String
    Dim str_sql As String
    Dim target_path As String
    
    
    Set fso = CreateObject("scripting.filesystemobject")
    
    source_path = Load.system_info.system_paths.stella_path & "stable_builds\stella_uw.accdb"
    source_path_placeholder_eur = Load.system_info.system_paths.stable_builds & "placeholder.accdb"
    source_path_placeholder_usa = Load.system_info.system_paths.stable_builds_us & "placeholder.accdb"
    
    'save current version to back-up folder
    target_path = Load.system_info.system_paths.old_stella_versions & Year(Date) & Month(Date) & Day(Date) & "-" & Hour(Time) & Minute(Time) & Second(Time) & " stella_uw.accdb"
    Call fso.CopyFile(source_path, target_path, True)
        
    'copy stable build to production area
    str_sql = "SELECT uw_id, uw_initials, budget_home_id, is_dev_id FROM " & sources.uws_view & " WHERE is_employed_id = " & menu_list.yes
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        Do Until rs.EOF = True
            
            If input_app_continent = Load.system_info.continents.americas Then
                target_path = Load.system_info.system_paths.individual_path_us & "Stella - " & !uw_initials & " " & rs!uw_id & ".accdb"
                Call fso.CopyFile(source_path, target_path, True)
                target_path = Load.system_info.system_paths.placeholder_path_published_us & "placeholder - " & rs!uw_id & ".accdb"
                Call fso.CopyFile(source_path_placeholder_usa, target_path, True)
            End If
            
            If input_app_continent = Load.system_info.continents.eurasia Then
                target_path = Load.system_info.system_paths.individual_path_eur & "Stella - " & !uw_initials & " " & rs!uw_id & ".accdb"
                Call fso.CopyFile(source_path, target_path, True)
                target_path = Load.system_info.system_paths.placeholder_path_published_eur & "placeholder - " & rs!uw_id & ".accdb"
                Call fso.CopyFile(source_path_placeholder_eur, target_path, True)
            End If
            
            .MoveNext
            
        Loop
        .Close
    End With
    Set rs = Nothing
    Set fso = Nothing
    
outro:
    utilities.call_stack_remove_last_item True
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "str_sql = " & str_sql, "", True
    Resume outro
End Sub
Public Sub update_support_apps(ByVal app_name As String)
    Load.call_stack = Load.call_stack & vbNewLine & "cls_form_system_dashboard.update_support_apps"
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim fso As Object
    Dim rs As ADODB.Recordset
    Dim source_path As String
    Dim str_sql As String
    Dim target_path As String
    
    Set fso = CreateObject("scripting.filesystemobject")
    
    source_path = Load.system_info.system_paths.stella_path & "stable_builds\" & app_name & ".accdb"
    
    'save current version to back-up folder
    target_path = Load.system_info.system_paths.old_stella_versions & Year(Date) & Month(Date) & Day(Date) & "-" & Hour(Time) & Minute(Time) & Second(Time) & " " & app_name & ".accdb"
    Call fso.CopyFile(source_path, target_path, True)
        
    'copy to American file server
    target_path = Load.system_info.system_paths.common_path_us & app_name & ".accdb"
    Call fso.CopyFile(source_path, target_path, True)

    'copy to European file server
    target_path = Load.system_info.system_paths.common_path & app_name & ".accdb"
    Call fso.CopyFile(source_path, target_path, True)

    Set fso = Nothing
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "cls_form_system_dashboard.update_support_apps"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With

End Sub

Public Sub default_values()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_form_system_dashboard.default_values"
    With Me
        .cmd_show_devs.field_value = False
        .cmd_filter_closed_issues.field_value = False
        .filter_routine_name.field_value = "_all"
        .filter_app_continent.field_value = "_all"
        .filter_app_name.field_value = "_all"
    End With
End Sub