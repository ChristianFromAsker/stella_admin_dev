Option Compare Database
Option Explicit
Public Function generate_buttons_for_deal_list()
    Dim buttons(20, 4) As Variant
    '0 = control name
    '1 = default caption
    '2 = control type
    '3 = caption for on
    '4 = caption for off
    
    Dim i As Integer
    i = 0
    buttons(i, 0) = 10
    i = i + 1
    buttons(i, 0) = "cmdNDA"
    buttons(i, 1) = "NDA: on"
    buttons(i, 2) = "status"
    buttons(i, 3) = "NDA: on"
    buttons(i, 4) = "NDA: off"
    i = i + 1
    buttons(i, 0) = "cmdNBI"
    buttons(i, 1) = "NBI: on"
    buttons(i, 2) = "status"
    buttons(i, 3) = "NBI: on"
    buttons(i, 4) = "NBI: off"
    i = i + 1
    buttons(i, 0) = "cmd_submission"
    buttons(i, 1) = "Sub: on"
    buttons(i, 2) = "status"
    buttons(i, 3) = "Sub: on"
    buttons(i, 4) = "Sub: off"
    i = i + 1
    buttons(i, 0) = "cmdPreferred"
    buttons(i, 1) = "Pref: on"
    buttons(i, 2) = "status"
    buttons(i, 3) = "Pref: on"
    buttons(i, 4) = "Pref: off"
    i = i + 1
    buttons(i, 0) = "cmdExpensed"
    buttons(i, 1) = "Expen: on"
    buttons(i, 2) = "status"
    buttons(i, 3) = "Expen: on"
    buttons(i, 4) = "Expen: off"
    i = i + 1
    buttons(i, 0) = "cmdUW"
    buttons(i, 1) = "UW: on"
    buttons(i, 2) = "status"
    buttons(i, 3) = "UW: on"
    buttons(i, 4) = "UW: off"
    i = i + 1
    buttons(i, 0) = "cmdSigned"
    buttons(i, 1) = "Signed: on"
    buttons(i, 2) = "status"
    buttons(i, 3) = "Signed: on"
    buttons(i, 4) = "Singed: off"
    i = i + 1
    buttons(i, 0) = "cmd_closed"
    buttons(i, 1) = "Closed: on"
    buttons(i, 2) = "status"
    buttons(i, 3) = "Closed: on"
    buttons(i, 4) = "Closed: off"
    i = i + 1
    buttons(i, 0) = "cmdLost"
    buttons(i, 1) = "Lost: off"
    buttons(i, 2) = "status"
    buttons(i, 3) = "Lost: on"
    buttons(i, 4) = "Lost: off"
    i = i + 1
    buttons(i, 0) = "cmdDeclined"
    buttons(i, 1) = "Decl: on"
    buttons(i, 2) = "status"
    buttons(i, 3) = "Decl: on"
    buttons(i, 4) = "Decl: off"
    i = i + 1
    buttons(i, 0) = "cmdCancel"
    buttons(i, 1) = "Cancel: on"
    buttons(i, 2) = "status"
    buttons(i, 3) = "Cancel: on"
    buttons(i, 4) = "Cancel: off"
    i = i + 1
    buttons(i, 0) = "cmd_collapsed"
    buttons(i, 1) = "Collaps: on"
    buttons(i, 2) = "status"
    buttons(i, 3) = "Collaps: on"
    buttons(i, 4) = "Collaps: off"
    i = i + 1

    generate_buttons_for_deal_list = buttons
    
End Function
Public Sub err_handler(ByVal input_proc_name As String _
, vba_err_no As Long _
, vba_err_desc As String _
, str_milestone As String _
, str_params As String _
, stella_err_desc As String _
, show_err As Boolean)
    Const proc_name As String = "central.err_handler"
    utilities.call_stack_add_item proc_name
    On Error Resume Next
    
    Dim cmd As ADODB.Command
    Dim app_name As String
    Dim app_continent As String
    
    If Load.is_debugging = True Then
        GoTo outro
    End If
    
    app_name = "-1"
    app_name = Load.system_info.app_name
    app_continent = "-1"
    app_continent = Load.system_info.app_continent
    
    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = Load.conn
        
        .CommandText = "INSERT INTO log_errors_t " & _
        "(system_error_text, system_error_code, stella_error_text, routine_name, call_stack, params, milestone, uw_name, app_name, file_path, app_continent) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        
        .CommandType = adCmdText

        ' Append parameters in the same order as the placeholders
        .Parameters.Append .CreateParameter("pSystemErrorText", adVarWChar, adParamInput, 255, vba_err_desc)
        .Parameters.Append .CreateParameter("pSystemErrorCode", adInteger, adParamInput, , vba_err_no)
        .Parameters.Append .CreateParameter("pStellaErrorText", adVarWChar, adParamInput, 255, stella_err_desc)
        .Parameters.Append .CreateParameter("pRoutineName", adVarWChar, adParamInput, 255, input_proc_name)
        .Parameters.Append .CreateParameter("pCallStack", adLongVarWChar, adParamInput, -1, Load.call_stack)
        .Parameters.Append .CreateParameter("pParams", adVarWChar, adParamInput, 255, str_params)
        .Parameters.Append .CreateParameter("pMilestone", adVarWChar, adParamInput, 255, str_milestone)
        .Parameters.Append .CreateParameter("pUwName", adVarWChar, adParamInput, 255, Environ("Username"))
        .Parameters.Append .CreateParameter("pAppName", adVarWChar, adParamInput, 255, Load.system_info.app_name)
        .Parameters.Append .CreateParameter("pFilePath", adVarWChar, adParamInput, 255, CurrentProject.FullName)
        .Parameters.Append .CreateParameter("pAppContinent", adVarWChar, adParamInput, 255, Load.system_info.app_continent)
        .Execute
    End With

    Set cmd = Nothing
    
    Err.Clear
    
    '16 October 2025, CK: I don't think the below is require anymore. _
    However, due to recent issues with this in the US, I am leaving it for now in case it is needed. But to be removed next time I read this :)
    
    If show_err = True Then
        MsgBox Load.system_info.error_instruction & vbNewLine & vbNewLine _
        & "Error description: " & vba_err_desc & vbNewLine _
        & "Where: " & input_proc_name & vbNewLine _
        & "Parameters: " & str_params & vbNewLine _
        & "App: " & Load.system_info.app_name _
        & vbNewLine & "Call stack: " & Right(call_stack, 500) _
        , , Load.system_info.error_msg_heading
    End If

outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Debug.Print "Error in " & proc_name & " - err.number: " & Err.Number & " - err.description: " & Err.Description
    Resume outro
End Sub
Public Sub open_external_resource_app(ByVal app_name As String _
, need_new_app As Boolean _
, app_file_path As String)
    Const proc_name As String = "central.open_external_resource_app"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim current_app As Access.Application
    Dim current_db_name As String
    Dim new_app As Access.Application
    Dim open_app As Boolean
    Dim str_milestone As String
    Dim str_app_path As String
    
    With Load.secondary_access_app
        
        str_app_path = app_file_path
        
        current_db_name = ""
        On Error Resume Next
        current_db_name = .CurrentDb.Name
        On Error GoTo err_handler
        If Load.is_debugging = True Then On Error GoTo 0
        
        open_app = False
        If current_db_name = "" Then
            open_app = True
        ElseIf current_db_name <> str_app_path Then
            .CloseCurrentDatabase
            open_app = True
        End If
        
        If open_app = True Then
            .OpenCurrentDatabase str_app_path, False
        End If
        
        str_milestone = "4"
        On Error GoTo err_path
        .Run "generate_test"
        .Visible = False
err_path:
        If Err.Number = 40351 Then
            .Visible = True
            Err.Clear
            MsgBox "PLEASE READ THIS. Complete the below instructions before moving on." _
            & vbNewLine & vbNewLine & "We have a small issue which I will need your help to solve." _
            & vbNewLine & vbNewLine & "Somwehere on your screen, there is a button which says 'enable content'. Please click it." _
            & vbNewLine & vbNewLine & "You may see it in the top left of your screen. It might appear as a small Access window. If so, just make it larger to show the button." _
            & vbNewLine & vbNewLine & "If do not see it, look on on your taskbar (the line at the bottom of your screen), for an Access (Stella) window called '" & app_name & "'." _
            & " Click on it to show the button." _
            & vbNewLine & vbNewLine & "If you don't see '" & app_name & "', you might need to hover your mouse cursor over the Access (Stella) app on the task bar. Then you'll see '" & app_name & "'." _
            & " Click on it to show the button." _
            & vbNewLine & vbNewLine & "When you have clicked on 'enable content', click 'ok' below." _
            , , "Small issue"
            
            On Error GoTo err_handler
            If Load.is_debugging = True Then On Error GoTo 0
        
            .CloseCurrentDatabase
            .Visible = False
            
            .OpenCurrentDatabase str_app_path, False
        End If
        
        On Error GoTo err_handler
        If Load.is_debugging = True Then On Error GoTo 0
    End With
    
outro:
    Exit Sub

err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, str_milestone, "", "", True
    Resume outro
End Sub

Public Sub update_deal_list_f(ByVal str_condition As String)
    'intro
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    Dim str_sql As String, rs As ADODB.Recordset, i As Integer
    
    'Open recordset
    With Forms("deal_list_f")
        str_sql = "SELECT * FROM " & Load.sources.global_deal_list_view & " WHERE " & str_condition
        Set rs = utilities.create_adodb_rs(conn, str_sql)
            If rs.EOF Then
                MsgBox "No deals to show", , "No deals"
                GoTo outro
            End If
            'Find total values for premium and limit
            rs.MoveFirst
            Dim total_premium As Currency
            Dim total_limit As Currency
            total_premium = 0
            total_limit = 0
            Do Until rs.EOF
                total_premium = total_premium + Nz(rs!total_rp_premium_on_deal_eur, 0)
                total_limit = total_limit + Nz(rs!total_rp_limit_on_deal_eur, 0)
                rs.MoveNext
            Loop
            .lblTotal.Caption = CLng(rs.RecordCount) & " hits"
            .total_premium = "Total premium: EUR " & Format(total_premium, "#,###,##0")
            .average_premium = "Average premium: EUR " & Format(total_premium / CLng(rs.RecordCount), "#,###,##0")
            .total_limit = "Total limit: EUR " & Format(total_limit, "#,###,##0")
            .average_limit = "Average limit: EUR " & Format(total_limit / CLng(rs.RecordCount), "#,###,##0")
            If total_limit = 0 Then
                .blended_rol = "Blended ROL: n/a "
            Else
                .blended_rol = "Blended ROL: " & Round(total_premium / total_limit, 4) * 100 & " %"
            End If
            
            'Decide height of form, and resize form
            Dim int_height As Long
            int_height = CLng(rs.RecordCount)
            Set .Recordset = rs
        rs.Close
        If int_height > 30 Then
            int_height = 20
        End If
        int_height = (300 * int_height) + 6500
        .SetFocus
        .InsideHeight = int_height
    End With
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: central.update_deal_list_f" & vbNewLine _
        & "Parameters: str_condition = " & str_condition & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub
Public Sub add_currencies_to_currencies_t()
    Load.check_conn_and_variables
    Dim app_access As Access.Application, app_access_path As String, rs As ADODB.Recordset, str_sql As String
    
    str_sql = "SELECT currency_id FROM " & sources.currencies_table _
    & " WHERE currency_date = '" & utilities.generate_sql_date(Date) & "'"
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        If CLng(.RecordCount) = 0 Then
            app_access_path = Load.system_info.system_paths.common_path & "currencies.accdb"
            Set app_access = CreateObject("Access.Application")
                With app_access
                    .OpenCurrentDatabase app_access_path, False
                    .Visible = False
                    .Run "feed_rates_to_db", utilities.generate_sql_date(Date)
                    .Quit
                End With
            Set app_access = Nothing
        End If
        .Close
    End With
End Sub
   
Public Sub stella_init()
    'CK 7 une 2023: legacy module. It's everywhere, so just redirecting to relevant sub rathern than replacing everywhere for now
    Load.init_global_variables
End Sub
Public Sub data_logger(ByVal log_object As cls_log_object, ByVal app_continent As String)
    Const proc_name As String = "central.data_logger"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim db_name As String
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim input_value As Variant
        
    If app_continent = Load.system_info.continents.eurasia Then
        db_name = "stella_eur."
    ElseIf app_continent = Load.system_info.continents.americas Then
        db_name = "stella_us."
    Else
        Debug.Print 1 / 0
    End If
    If log_object.new_value_var <> "" And log_object.new_value_number <> "" Then
        MsgBox "Please snip this message to Christian asap." & vbNewLine & vbNewLine _
            & "Error description: Both log_object.new_value_var and log_object.new_value_number had values." & vbNewLine _
            & "Where: central.data_logger" & vbNewLine _
            & "Parameters: log_object.data_set_id = " & log_object.data_set_id & ", log_object.record_id = " & log_object.record_id & vbNewLine _
            & "App: Stella UW Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
        GoTo outro
    End If
                
    If log_object.new_value_var <> "" Then
        str_sql = "INSERT INTO " & db_name & sources.data_log_table & " (data_set_id, field_name, new_value_text, changer_id, comment, record_id) VALUES(" _
            & log_object.data_set_id _
            & ", '" & log_object.field_name & "'" _
            & ", '" & log_object.new_value_var & "'" _
            & ", " & log_object.changer_id _
            & ", '" & log_object.comment & "'" _
            & ", " & log_object.record_id _
            & ")"
    End If
    
    If log_object.new_value_number <> "" Then
        str_sql = "INSERT INTO " & db_name & sources.data_log_table & " (data_set_id, field_name, new_value_number, changer_id, comment, record_id) VALUES(" _
            & log_object.data_set_id _
            & ", '" & log_object.field_name & "'" _
            & ", " & log_object.new_value_number _
            & ", " & log_object.changer_id _
            & ", '" & log_object.comment & "'" _
            & ", " & log_object.record_id _
            & ")"
    End If
    
    If log_object.executed_sql <> "" Then
        str_sql = "INSERT INTO " & db_name & sources.data_log_table & " (data_set_id, executed_sql, changer_id, comment, record_id) VALUES(" _
            & log_object.data_set_id _
            & ", '" & Replace(log_object.executed_sql, "'", "''") & "'" _
            & ", " & log_object.changer_id _
            & ", '" & log_object.comment & "'" _
            & ", " & log_object.record_id _
            & ")"
    End If
    
    conn.Execute str_sql
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: central.data_logger" & vbNewLine _
        & "Parameters: log_object.data_set_id = " & log_object.data_set_id & ", log_object.record_id = " & log_object.record_id & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub
Public Sub auto_archiving()
    'Purpose: Set status for all risks to lost if deal is old and has not been progressed
    'intro
    On Error GoTo err_handler
    If Load.is_init = False Then Central.stella_init
    
    Dim rs As ADODB.Recordset
    Dim log_object As cls_log_object, str_sql As String
    
    'set old deals with status nda to cancelled
    str_sql = "SELECT deal_id FROM stella_eur." & Load.sources.deals_view & " WHERE " _
        & " deal_status_id = " & deal_statuses.nda _
        & " AND (DATEDIFF(CURRENT_DATE(), create_date) / 30 - months_before_auto_archiving) > 0"
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        If rs.BOF And rs.EOF Then
        
        Else
            Do Until rs.EOF = True
                Set log_object = New cls_log_object
                With log_object
                    .changer_id = Load.stella_uw_id
                    .data_set_id = sources.deals_table_id
                    .field_name = "deal_status_id"
                    .new_value_number = deal_statuses.cancelled
                    .record_id = rs!deal_id
                End With
                Central.data_logger log_object, Load.system_info.continents.eurasia
                Set log_object = Nothing
                conn.Execute "UPDATE " & sources.deals_table & " SET deal_status_id = " & deal_statuses.cancelled & " WHERE deal_id = " & rs!deal_id
                rs.MoveNext
            Loop
        End If
    rs.Close
    
    'set old deals with status nbi or submission to lost
    str_sql = "SELECT deal_id FROM stella_eur." & sources.deals_for_auto_archiving_view
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        If rs.BOF And rs.EOF Then
            GoTo outro
        End If
        Do Until rs.EOF = True
            Set log_object = New cls_log_object
            With log_object
                .changer_id = Load.stella_uw_id
                .data_set_id = sources.deals_table_id
                .field_name = "deal_status_id"
                .new_value_number = deal_statuses.lost
                .record_id = rs!deal_id
            End With
            Central.data_logger log_object, Load.system_info.continents.eurasia
            Set log_object = Nothing
            conn.Execute "UPDATE " & sources.deals_table & " SET deal_status_id = " & deal_statuses.lost & " WHERE deal_id = " & rs!deal_id
            rs.MoveNext
        Loop
    rs.Close
    
outro:
    Set rs = Nothing
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: central.auto_archiving" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"

End Sub
Public Function generate_deal_object(ByVal deal_id As Long) As Object
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim obj_deal As cls_deal
    Set obj_deal = New cls_deal
    ' Populate the deal object
    obj_deal.deal_id = deal_id
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    str_sql = "SELECT * FROM " & Load.sources.deals_view & " WHERE deal_id = " & obj_deal.deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        rs.MoveFirst
        With obj_deal
            .create_date = rs!create_date
            .deal_id = deal_id
            If IsNull(rs!deal_name) = False Then
                .deal_name = rs!deal_name
            Else
                MsgBox "You need to create a deal folder before you can open it", , "No folder exists"
                Exit Function
            End If
            .broker_firm = rs!broker_firm_id
            .broker_firm_hr = Nz(rs!broker_firm, "-1")
            .spa_law = rs!spa_law
            If Not IsNull(rs!inception_date) Then
                .inception_date = rs!inception_date
            End If
            .status = rs!deal_status_id
        End With
    rs.Close
    Set rs = Nothing
    Set generate_deal_object = obj_deal
    
outro:
    Exit Function
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "central.generate_deal_object"
        .milestone = "str_sql = " & str_sql
        .params = "deal_id = " & deal_id
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
End Function
Private Function create_date_string()
Dim str_year As String, str_month As String, str_day As String
str_year = Year(Date)
If Month(Date) > 9 Then
    str_month = Month(Date)
Else
    str_month = "0" & Month(Date)
End If
If Day(Date) > 9 Then
    str_day = Day(Date)
Else
    str_day = "0" & Day(Date)
End If
create_date_string = str_year & "." & str_month & "." & str_day
End Function
Public Sub reset_variables()
    Load.is_init = False
    Set conn = Nothing
    Central.stella_init
End Sub