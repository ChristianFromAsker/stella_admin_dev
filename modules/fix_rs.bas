Option Compare Database
Option Explicit
Public Sub deal_log_f(ByVal deal_id As Long)
    Const proc_name As String = "fix_rs.deal_log_f"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim rs As ADODB.Recordset
    Dim str_form As String
    Dim str_sql As String
        
    str_form = "deal_log_f"
    
    str_sql = "SELECT" _
    & " l.record_id   data_log_id" & vbNewLine _
    & ", l.app_name" & vbNewLine _
    & ", l.change_source" & vbNewLine _
    & ", l.comment" & vbNewLine _
    & ", l.data_set_id  data_set" & vbNewLine _
    & ", l.field_name" & vbNewLine _
    & " , IF(" _
        & "l.new_value IS NULL, l.new_value_text, l.new_value" _
    & ")   new_value" & vbNewLine _
    & " , l.record_id" & vbNewLine _
    & " , l.executed_sql" & vbNewLine _
    & " , l.changer_id      changer" & vbNewLine _
    & " , l.create_date" & vbNewLine _
    & " , r.deal_id         deal_id_referral" & vbNewLine _
    & " , p.deal_id         deal_id_policy" & vbNewLine _
    & " , d.deal_name"
    
    str_sql = str_sql & " FROM" & vbNewLine _
    & " stella_eur.log_data_t l" & vbNewLine _
    & " LEFT JOIN" & vbNewLine _
        & " stella_eur.cm_referrals_t r" & vbNewLine _
        & " ON r.id = l.record_id AND data_set_id LIKE '%cm_referrals_t'" & vbNewLine _
    & " LEFT JOIN" _
    & "     stella_eur.deals_t d" _
    & "     ON d.deal_id = l.record_id AND data_set_id = 'deals_t'" _
    & " LEFT JOIN" _
        & " stella_eur.cm_deal_questions_t dq" _
        & " ON dq.id = l.record_id AND data_set_id LIKE '%cm_deal_questions_t'" _
    & " LEFT JOIN" _
        & " stella_eur.layers_t p" _
        & " ON p.id = l.record_id AND l.data_set_id LIKE '%layers_t'" _
    & " WHERE" _
    & " (" _
        & " l.record_id = " & deal_id _
        & " AND data_set_id IN ('deals_t', 'stella_eur.cm_deals_t')" _
    & ")" _
        & " OR p.deal_id = " & deal_id _
        & " OR r.deal_id = " & deal_id _
        & " OR dq.deal_id = " & deal_id _
    & " ORDER BY l.create_date DESC"

    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        Set Forms(str_form).Recordset = rs
        .Close
    End With
    Set rs = Nothing
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub system_dashboard_f()
    Load.call_stack = Load.call_stack & vbNewLine & "fix_rs.system_dashboard_f"
    Dim str_form As String, str_sql As String, rs As ADODB.Recordset
    
    str_form = global_vars.system_dashboard.form_name
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    Dim str_where As String
    str_where = ""
    With global_vars.system_dashboard
        If .cmd_show_devs.field_value = False Then
            str_where = " AND (NOT uw_name = 'christian.kartnes' AND NOT uw_name = 'tom.evans')"
        End If
        
        If .filter_routine_name.field_value <> "_all" Then
            str_where = str_where & " AND routine_name = '" & .filter_routine_name.field_value & "'"
        End If
        
        If .filter_app_name.field_value <> "_all" Then
            str_where = str_where & " AND app_name = '" & .filter_app_name.field_value & "'"
        End If
'        If .filter_app_continent.field_value <> "_all" Then
'            str_where = str_where & " AND app_continent = '" & .filter_app_continent.field_value & "'"
'        End If
        If .cmd_filter_closed_issues.field_value = False Then
            str_where = str_where & " AND is_closed_id = 94"
        End If
    End With
    str_sql = "SELECT * FROM " & Load.sources.log_errors_v & " WHERE 1" & str_where & "  ORDER BY create_date DESC"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        Set Forms(str_form).Recordset = rs
        .Close
    End With
    Set rs = Nothing
    
outro:
    Exit Sub

End Sub
Public Sub awac_excels_f()
    Dim str_form As String, str_sql As String, rs As ADODB.Recordset
    str_form = "awac_excels_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    str_sql = "SELECT * FROM " & Load.sources.awac_excels_view & " ORDER BY email_address"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        Set Forms(str_form).Recordset = rs
        .Close
    End With
    Set rs = Nothing
    
outro:
    Exit Sub
    
End Sub
Public Sub service_messages_f(Optional ByVal input_condition As String)
    Dim str_form As String
    str_form = "service_messages_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    Dim str_sql As String, str_where As String
    str_where = ""
    If input_condition <> "" Then str_where = " WHERE " & input_condition & " "
        
    str_sql = "SELECT * FROM " & Load.sources.service_messages_view & str_where & " ORDER BY message_id DESC"
    Dim rs As ADODB.Recordset
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        Set Forms(str_form).Recordset = rs
        .Close
    End With
outro:
    Set rs = Nothing
    Exit Sub
    
End Sub
Public Sub global_policy_list_f()
    Const proc_name As String = "fix_rs.global_policy_list_f"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim date_end As String
    Dim date_start As String
    Dim rp_entity As cls_field
    Dim rs As ADODB.Recordset
    Dim str_form As String
    Dim str_sql As String
    Dim str_sql_entity As String
    
    str_form = "global_policy_list_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    
    str_sql = "SELECT * FROM " & Load.sources.global_policy_list_view & " WHERE (deal_status_id = 6 OR deal_status_id = 436)"
    
    'RP entity condition
    str_sql_entity = ""
    For Each rp_entity In global_policy_list.col_navins_homes
        If rp_entity.field_value = True Then
            If str_sql_entity <> "" Then str_sql_entity = str_sql_entity & " OR "
            str_sql_entity = str_sql_entity & "issuing_entity_id = " & rp_entity.id
        End If
    Next rp_entity
    
    If str_sql_entity = "" Then
        MsgBox "You must choose at least one RP Entity.", , "No RP Entity"
        GoTo outro
    End If
    
    str_sql = str_sql & " AND (" & str_sql_entity & ")"
    
    'date condition
    With Forms(str_form)
        If IsNull(!date_month_start) = False And IsNull(!date_month_end) = False Then
            date_start = utilities.generate_sql_date("1-" & !date_month_start & "-" & !date_year_start)
            date_end = utilities.generate_sql_date("31-" & !date_month_end & "-" & !date_year_end)
            str_sql = str_sql & " AND (inception_date >= '" & date_start & "' AND inception_date <= '" & date_end & "')"
        End If
    End With
    
    'name or policy no condition
    With Forms(str_form)
        If IsEmpty(!deal_name_search) = False And !deal_name_search <> "" Then
            str_sql = str_sql & " AND (policy_no LIKE '%" & !deal_name_search & "%' OR deal_name LIKE '%" & !deal_name_search & "%')"
        End If
    End With
    
    'set sort condition
    str_sql = str_sql & " ORDER BY " & global_policy_list.sort_field & " " & global_policy_list.sort_order
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        Set Forms(str_form).Recordset = rs
        Dim detail_height As Long
        detail_height = 6000 + utilities.twips_converter(0.6, "cm") * CLng(rs.RecordCount)
        If detail_height > 10000 Then detail_height = 12000
        With Forms(str_form)
            .InsideHeight = detail_height
            !lbl_policy_count.Caption = "policy count: " & CLng(rs.RecordCount)
        End With
    rs.Close
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "str_sql = " & str_sql, "", "", True
    Resume outro
End Sub
Public Sub pricing_guidelines_f(Optional ByVal jurisdiction_id As Long, Optional ByVal super_sector_id As Long)
    'latest review or change: 19 October 2023
    'intro
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    If Load.is_debugging = True Then On Error GoTo 0
    Dim str_form As String
    str_form = "pricing_guidelines_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    
    Dim str_sql As String, rs As ADODB.Recordset
    str_sql = "SELECT * FROM " & Load.sources.pricing_guidelines_view & " WHERE is_active = 93"
    If jurisdiction_id <> 0 Then str_sql = str_sql & " AND jurisdiction_id = " & jurisdiction_id
    If super_sector_id <> 0 Then str_sql = str_sql & " AND super_sector_id = " & super_sector_id
    str_sql = str_sql & " ORDER BY sector_name"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        Set Forms(str_form).Recordset = rs
        Dim detail_height As Long
        detail_height = 2500 + utilities.twips_converter(0.6, "cm") * CLng(rs.RecordCount)
        If detail_height > 10000 Then detail_height = 12000
        Forms(str_form).InsideHeight = detail_height
    rs.Close
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: fix_rs.pricing_guidelines_f" & vbNewLine _
        & "Parameters: str_sql = " & str_sql & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub
Public Sub brands_f(ByVal include_inactive As Boolean)
    Const proc_name As String = "fix_rs.brands_f"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String
    Dim str_form As String
    Dim rs As ADODB.Recordset
    
    str_sql = "SELECT * FROM " & sources.brands_table & " WHERE is_deleted = 0"
    If include_inactive = False Then str_sql = str_sql & " AND is_active = 93"
    str_sql = str_sql & " ORDER BY brand_name"
    
    str_form = "brands_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        Set Forms(str_form).Recordset = rs
    rs.Close
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "fix_rs.brands_f"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub broker_firms_f()
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0

    Dim str_form As String, str_sql As String, rs As ADODB.Recordset
    str_form = "broker_firms_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    str_sql = "SELECT * FROM " & sources.broker_firms_view & " ORDER BY business_name"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        Set Forms(str_form).Recordset = rs
    rs.Close
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "fix_rs.broker_firm_f"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
    GoTo outro
End Sub
Public Sub broker_persons_f()
    Dim str_form As String, str_sql As String, rs As ADODB.Recordset
    str_form = "broker_person_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    str_sql = "SELECT * FROM broker_persons_v ORDER BY personal_name"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        Set Forms(str_form).Recordset = rs
    rs.Close
    
outro:
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .call_stack = Load.call_stack
        .routine_name = "fix_rs.broker_persons_f"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub

Public Sub law_firms_f()
    Dim str_form As String, str_sql As String, rs As ADODB.Recordset
    str_form = "law_firms_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    str_sql = "SELECT * FROM " & sources.law_firms_view & " ORDER BY firm_name"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        Set Forms(str_form).Recordset = rs
    rs.Close
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: fix_rs.law_firms_f" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub

Public Sub lawyers_f()

    Dim str_form As String
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    
    str_form = "lawyers_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    str_sql = "SELECT * FROM lawyers_v ORDER BY personal_name"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        Set Forms(str_form).Recordset = rs
    rs.Close
outro:
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: fix_rs.add_lawyer" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub
Public Sub uws_f(ByVal str_condition As String, _
    ByVal str_source As String _
)
    Const proc_name As String = "fix_rs.uws_f"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0

    Dim rs As ADODB.Recordset
    Dim str_sql As String
    Dim str_form As String
    
    str_form = "uws_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    
    str_sql = "SELECT * from " & str_source & " WHERE 1 " & str_condition & " ORDER BY uw_name ASC, employee_role_start_date DESC"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        Set Forms(str_form).Recordset = rs
    rs.Close
    Set rs = Nothing
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "str_sql = " & str_sql, "", "", True
    Resume outro
End Sub
Public Sub uws_f_refresh()
    Const proc_name As String = "fix_rs.uws_f_refresh"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0

    Const str_form As String = "uws_f"

    Dim rs As ADODB.Recordset
    Dim str_sql As String
    
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    
    str_sql = Forms(str_form).RecordSource
    Set rs = utilities.create_adodb_rs(Load.conn, str_sql)
        Set Forms(str_form).Recordset = rs
    rs.Close
    Set rs = Nothing
    
outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "str_sql = " & str_sql, "", "", True
    Resume outro
End Sub
Public Sub jurisdictions_f()
    'intro
    On Error GoTo err_handler

    Dim str_form As String
    str_form = "jurisdictions_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        Exit Sub
    End If
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    str_sql = "SELECT * FROM " & sources.jurisdictions_view & " WHERE " _
        & "(jurisdiction_type = 'country' OR jurisdiction_type = 'us_state') " _
        & " ORDER BY jurisdiction_type, jurisdiction"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        Set Forms(str_form).Recordset = rs
    rs.Close
    Set rs = Nothing
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: fix_rs.add_country_f" & vbNewLine _
        & "Parameters: str_sql = " & str_sql & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
End Sub
Public Sub financial_advisors_f()
    On Error GoTo err_handler

    Dim str_form As String, str_sql As String, rs As ADODB.Recordset
    str_form = "add_financial_advisor_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    str_sql = "SELECT * FROM " & sources.financial_advisors_view & " ORDER BY firm_name"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        Set Forms(str_form).Recordset = rs
    rs.Close
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: fix_rs.broker_firms_f" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub
Public Sub insurers_f()
    'intro
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0

    Dim str_form As String, str_sql As String, rs As ADODB.Recordset
    str_form = "add_insurer_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    str_sql = "SELECT * FROM " & sources.insurers_view & " ORDER BY insurer_business_name"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        Set Forms(str_form).Recordset = rs
    rs.Close

outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
    
err_handler:
    MsgBox "Something went wrong. Snip this to Tom and Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: fix_rs.insurers_f" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub
Public Sub templates_f()
    'latest review or change: 7 Jun 2023 by CK
    'intro
    On Error GoTo err_handler

    Dim str_form As String, str_sql As String, rs As ADODB.Recordset
    str_form = "add_templates_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        GoTo outro
    End If
    'str_sql = "SELECT * FROM " & sources.templates_table & " WHERE is_deleted = 0 ORDER BY template_name"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        Set Forms(str_form).Recordset = rs
    rs.Close
    
outro:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: fix_rs.insurers_f" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub