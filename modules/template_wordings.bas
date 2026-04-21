Option Compare Database

Public Sub new_sub()
    Load.call_stack = Load.call_stack & vbNewLine & "tempalte_wordings.new_sub"
    Load.check_conn_and_variables
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "tempalte_wordings.new_sub"
        .milestone = ""
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub