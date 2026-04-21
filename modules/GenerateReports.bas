Option Compare Database
Option Explicit
Public Sub generate_master_list_extract(str_sql As String)
    Dim proc_name As String
    proc_name = "GenerateReports.generate_master_list_extract"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_milestone As String
    Dim err_sql As String
    Dim app_access As Access.Application
    Dim app_access_path As String
    
    Load.check_secondary_access_app
    With Load.secondary_access_app
        On Error Resume Next
        .CloseCurrentDatabase
        On Error GoTo err_handler
        If Load.is_debugging = True Then On Error GoTo 0

        .OpenCurrentDatabase Load.system_info.system_paths.common_path & "reports.accdb", False
        .Visible = False
        .Run "reports.generate_eur_us_deals_report", str_sql
        .CloseCurrentDatabase
        .OpenCurrentDatabase Load.system_info.system_paths.common_path & "placeholder.accdb", False
    End With

outro:
    Exit Sub

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Sub
Public Sub global_policy_extract(str_sql As String)
    Dim proc_name As String
    proc_name = "GenerateReports.global_policy_extract"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_milestone As String
    Dim err_sql As String
    Dim app_access As Access.Application
    Dim app_access_path As String
    
    Load.check_secondary_access_app
    With Load.secondary_access_app
        On Error Resume Next
        .CloseCurrentDatabase
        On Error GoTo err_handler
        If Load.is_debugging = True Then On Error GoTo 0

        .OpenCurrentDatabase Load.system_info.system_paths.common_path & "reports.accdb", False
        .Visible = False
        .Run "reports.generate_global_policy_report", str_sql
        .CloseCurrentDatabase
        .OpenCurrentDatabase Load.system_info.system_paths.common_path & "placeholder.accdb", False
    End With

outro:
    Exit Sub

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub