Option Compare Database
Option Explicit

Public Type typ_header_controls
    header_chooser_year As New cls_field
End Type

Public Type typ_uw_data_controls
    uw_name As New cls_field
    nickname As New cls_field
    uw_initials As New cls_field
    email As New cls_field
    employee_role_id As New cls_field
    is_employed_id As New cls_field
    budget_home_id As New cls_field
    date_start As New cls_field
    date_end As New cls_field
    usertype As New cls_field
    has_admin_access_id As New cls_field
    can_change_general_id As New cls_field
    can_change_jurisdictions_id As New cls_field
    is_dev_id As New cls_field
    can_change_uws_id As New cls_field
    can_change_budget_home_id As New cls_field
    is_regional_lead_id As New cls_field
    uw_id As New cls_field
    username As New cls_field
End Type

Public col_uw_data_controls As Collection
Public current_source As String
Public header_controls As user_management.typ_header_controls
Public uw_data_controls As user_management.typ_uw_data_controls

Public Sub uw_statistics(ByVal str_condition As String)
    Const proc_name As String = "user_management.uw_statistics"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Const form_name As String = "uws_f"
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    str_sql = "SELECT * FROM " & Load.sources.uw_roles_view & " WHERE 1 " & str_condition
    
    Set rs = utilities.create_adodb_rs(Load.conn, str_sql)
        Forms(form_name)!role_count = rs.RecordCount
    rs.Close
    Set rs = Nothing
    
    ' if forms(form_name").controls(
    str_sql = "SELECT * FROM " & Load.sources.uw_role_fte_by_year_view & " WHERE "
    
outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub populate_uws_role(ByVal employee_role_link_id As Long, ByVal new_role As Boolean)
    Const proc_name As String = "user_management.populate_uws_add_f_with_existing_user"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Const str_form As String = "uws_role_f"
    
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    str_sql = "SELECT  employee_role_link_id, uw_id, role_id, employee_role_start_date, employee_role_end_date, uw_name" _
    & " FROM " & Load.sources.uw_roles_view _
    & " WHERE employee_role_link_id = " & employee_role_link_id
    
    Set rs = utilities.create_adodb_rs(Load.conn, str_sql)
        With Forms(str_form)
            !employee_role_link_id = rs!employee_role_link_id
            !uw_id__underwriters_t = rs!uw_id
            !uw_name = rs!uw_name
            If new_role = False Then
                !employee_role_id__employee_roles_t = rs!role_id
                !employee_role_start_date = rs!employee_role_start_date
                !employee_role_end_date = rs!employee_role_end_date
            End If
        End With
    rs.Close
    Set rs = Nothing
outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "employee_role_link_id = " & employee_role_link_id, "", True
    Resume outro
End Sub

Public Sub populate_uws_add_f_with_existing_user(ByVal uw_id As Long)
    Const proc_name As String = "user_management.populate_uws_add_f_with_existing_user"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Const str_form As String = "uws_add_f"
    
    Dim form_field As cls_field
    Dim col_fields As Collection
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    Set col_fields = New Collection
    With user_management.uw_data_controls
        col_fields.Add .budget_home_id
        col_fields.Add .can_change_budget_home_id
        col_fields.Add .can_change_general_id
        col_fields.Add .can_change_jurisdictions_id
        col_fields.Add .can_change_uws_id
        col_fields.Add .date_end
        col_fields.Add .date_start
        col_fields.Add .email
        col_fields.Add .has_admin_access_id
        col_fields.Add .is_dev_id
        col_fields.Add .is_employed_id
        col_fields.Add .is_regional_lead_id
        col_fields.Add .nickname
        col_fields.Add .username
        col_fields.Add .usertype
        col_fields.Add .uw_id
        col_fields.Add .uw_initials
        col_fields.Add .uw_name
    End With
    
    str_sql = "SELECT uw_id, budget_home_id, can_change_general_id, can_change_budget_home_id, can_change_jurisdictions_id, can_change_uws_id" _
    & ", email, end_date, has_admin_access_id" _
    & ", is_dev_id, is_employed_id, is_regional_lead_id" _
    & ", nickname, start_date, uw_initials, uw_name" _
    & ", user_type_id, user_name " _
    & " FROM " & Load.sources.uws_table _
    & " WHERE uw_id = " & uw_id
    
    Set rs = utilities.create_adodb_rs(Load.conn, str_sql)
    For Each form_field In col_fields
        With form_field
            Forms(str_form).Controls(.field_name) = rs.Fields(.field_name_in_table)
        End With
    Next form_field
    rs.Close
    Set rs = Nothing
outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "uw_id = " & uw_id, "", True
    Resume outro
End Sub
Public Sub copy_cm_and_stella(ByVal uw_id As Long, ByVal str_initials As String)
    Const proc_name As String = "user_management.copy_cm_and_stella"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim fso As Object
    Dim target_path_cm As String
    Dim target_path_stella As String
    Dim target_path_placeholder As String
    Dim source_path_cm As String
    Dim source_path_stella As String
    Dim source_path_placeholder As String
    
    source_path_cm = Load.system_info.system_paths.stable_builds & "cm_uw.accdb"
    source_path_stella = Load.system_info.system_paths.stable_builds & "stella_uw.accdb"
    source_path_placeholder = Load.system_info.system_paths.stable_builds & "placeholder.accdb"
    
    target_path_stella = Load.system_info.system_paths.stella_path & "published\individual\stella - " & str_initials & " " & uw_id & ".accdb"
    target_path_cm = Load.system_info.system_paths.stella_path & "published\individual\cm - " & str_initials & " " & uw_id & ".accdb"
    target_path_placeholder = Load.system_info.system_paths.published_eur & "placeholders\placeholder - " & uw_id & ".accdb"
            
    Set fso = CreateObject("scripting.filesystemobject")
    fso.CopyFile source_path_cm, target_path_cm
    fso.CopyFile source_path_stella, target_path_stella
    fso.CopyFile source_path_placeholder, target_path_placeholder
outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "uw_id = " & uw_id, "", True
    Resume outro
End Sub
Public Sub init()
    Const proc_name As String = "user_management.init"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    With user_management.header_controls.header_chooser_year
        .field_name = "header_chooser_year"
        .field_visible = True
    End With
    
    user_management.init__cols
    user_management.init__uw_data_controls
outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub init__cols()
    Const proc_name As String = "user_management.init__cols"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Set user_management.col_uw_data_controls = New Collection
    With user_management.uw_data_controls
        user_management.col_uw_data_controls.Add .budget_home_id
        user_management.col_uw_data_controls.Add .can_change_budget_home_id
        user_management.col_uw_data_controls.Add .can_change_general_id
        user_management.col_uw_data_controls.Add .can_change_jurisdictions_id
        user_management.col_uw_data_controls.Add .can_change_uws_id
        user_management.col_uw_data_controls.Add .date_end
        user_management.col_uw_data_controls.Add .date_start
        user_management.col_uw_data_controls.Add .email
        user_management.col_uw_data_controls.Add .employee_role_id
        user_management.col_uw_data_controls.Add .has_admin_access_id
        user_management.col_uw_data_controls.Add .is_dev_id
        user_management.col_uw_data_controls.Add .is_employed_id
        user_management.col_uw_data_controls.Add .is_regional_lead_id
        user_management.col_uw_data_controls.Add .nickname
        user_management.col_uw_data_controls.Add .username
        user_management.col_uw_data_controls.Add .usertype
        user_management.col_uw_data_controls.Add .uw_id
        user_management.col_uw_data_controls.Add .uw_initials
        user_management.col_uw_data_controls.Add .uw_name
    End With
outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Function does_user_exist(ByVal username As String) As Boolean
    Const proc_name As String = "user_management.does_user_exist"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim field_name As String
    Dim output As Boolean
    Dim rs As ADODB.Recordset
    Dim str_sql As String
            
    output = False
    field_name = user_management.uw_data_controls.username.field_name_in_table
    str_sql = "SELECT uw_id FROM " & Load.sources.uws_table & " WHERE " & field_name & " = '" & username & "'"
    Set rs = utilities.create_adodb_rs(Load.conn, str_sql)
        If rs.RecordCount > 0 Then output = True
    rs.Close
    Set rs = Nothing
    
    does_user_exist = output
outro:
    utilities.call_stack_remove_last_item False
    Exit Function
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Function

Public Sub init__uw_data_controls()
    Const proc_name As String = "user_management.init__uw_data_controls"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim form_field As cls_field
    For Each form_field In user_management.col_uw_data_controls
        form_field.field_bg_color = Load.colors.white
    Next form_field
    
    With user_management.uw_data_controls
        With .budget_home_id
            .field_name = "budget_home_id"
            .field_name_in_table = "budget_home_id"
            .field_visible = True
            .is_mandatory = True
        End With
        With .can_change_budget_home_id
            .field_name = "can_change_budget_home_id"
            .field_name_in_table = "can_change_budget_home_id"
            .field_visible = True
            .is_mandatory = True
        End With
        With .can_change_general_id
            .field_name = "can_change_general_id"
            .field_name_in_table = "can_change_general_id"
            .field_visible = True
        End With
        With .can_change_jurisdictions_id
            .field_name = "can_change_jurisdictions_id"
            .field_name_in_table = "can_change_jurisdictions_id"
            .field_visible = True
        End With
        With .can_change_uws_id
            .field_name = "can_change_uws_id"
            .field_name_in_table = "can_change_uws_id"
            .field_visible = True
        End With
        With .date_end
            .field_name = "date_end"
            .field_name_in_table = "end_date"
            .field_visible = True
            .is_mandatory = False
        End With
        With .date_start
            .field_name = "date_start"
            .field_name_in_table = "start_date"
            .field_visible = True
            .is_mandatory = True
        End With
        With .email
            .field_name = "email"
            .field_name_in_table = "email"
            .field_visible = True
            .is_mandatory = True
        End With
        With .employee_role_id
            .field_name = "employee_role_id"
            .field_name_in_table = "employee_role_id"
            .is_mandatory = True
        End With
        With .has_admin_access_id
            .field_name = "has_admin_access_id"
            .field_name_in_table = "has_admin_access_id"
            .field_visible = True
        End With
        With .is_dev_id
            .field_name = "is_dev_id"
            .field_name_in_table = "is_dev_id"
            .field_visible = True
        End With
        With .is_employed_id
            .field_name = "is_employed_id"
            .field_name_in_table = "is_employed_id"
            .field_visible = True
        End With
        With .is_regional_lead_id
            .field_name = "is_regional_lead_id"
            .field_name_in_table = "is_regional_lead_id"
            .field_visible = True
        End With
        With .nickname
            .field_name = "nickname"
            .field_name_in_table = "nickname"
            .field_visible = True
            .is_mandatory = True
        End With
        With .username
            .field_name = "username"
            .field_name_in_table = "user_name"
            .field_visible = True
        End With
        With .usertype
            .field_name = "usertype_id"
            .field_name_in_table = "user_type_id"
            .field_visible = True
            .is_mandatory = True
        End With
        With .uw_id
            .field_name = "uw_id"
            .field_name_in_table = "uw_id"
            .field_visible = True
        End With
        With .uw_initials
            .field_name = "uw_initials"
            .field_name_in_table = "uw_initials"
            .field_visible = True
            .is_mandatory = True
        End With
        With .uw_name
            .field_name = "uw_name"
            .field_name_in_table = "uw_name"
            .field_visible = True
            .is_mandatory = True
        End With
    End With
    
outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Function add_new_uw_id() As Long
    Const proc_name As String = "user_management.add_new_uw_id"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    str_sql = "INSERT INTO " & Load.sources.uws_table & " (is_deleted) VALUES (1)"
    conn.Execute str_sql
    str_sql = "SELECT uw_id FROM " & Load.sources.uws_table & " WHERE is_deleted = 1 ORDER BY uw_id DESC LIMIT 1"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        add_new_uw_id = !uw_id.Value
        .Close
    End With
    Set rs = Nothing
outro:
    utilities.call_stack_remove_last_item False
    Exit Function
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Function