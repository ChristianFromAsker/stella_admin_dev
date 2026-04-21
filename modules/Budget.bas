Option Compare Database
Option Explicit
Public wsBudget As Excel.Worksheet
Public wbBudget As Excel.Workbook
Public Sub call_global_budget(ByVal lon_year As Long)
    Dim proc_name As String
    proc_name = "Budget.global_budget"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.check_secondary_access_app
    With Load.secondary_access_app
        On Error Resume Next
        .CloseCurrentDatabase
        On Error GoTo err_handler
        If Load.is_debugging = True Then On Error GoTo 0

        .OpenCurrentDatabase Load.system_info.system_paths.common_path & "reports.accdb", False
        
        .Visible = False
        If Load.is_debugging = True Then .Visible = True
            
        .Run "reports.budget_entry", lon_year
        .CloseCurrentDatabase
        .OpenCurrentDatabase Load.system_info.system_paths.common_path & "placeholder.accdb", False
    End With
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "Budget.global_budget"
        .milestone = "app_access_path = " & Load.system_info.system_paths.reports_app
        .params = "int_year = " & lon_year
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
End Sub