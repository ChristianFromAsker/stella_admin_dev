Option Compare Database
Option Explicit

Public Sub deal_log()
    Const proc_name As String = "open_forms.deal_log"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String
    
    str_form = "deal_log_f"
    
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    DoCmd.MoveSize Right:=200, Down:=200, Width:=28000, Height:=12000
    fix_rs.deal_log_f 0
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub

err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub jurisdictions_f()
    Load.check_conn_and_variables
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    Dim str_form As String
    str_form = "jurisdictions_f"
    
    Dim i As Integer
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    Dim arr_controls() As Variant
    ReDim arr_controls(0 To 3, 0 To 1)
    
    i = 0
    arr_controls(i, 0) = "new_budget_home"
    arr_controls(i, 1) = "SELECT entity_id id, entity_business_name menu_item FROM " & Load.sources.entities_table _
    & " WHERE entity_type = 475 AND is_deleted = 0 ORDER BY entity_business_name"
    
    i = i + 1
    arr_controls(i, 0) = "new_working_folder"
    arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & Load.sources.menu_list_table & " WHERE is_deleted = 0 AND item_type = 'WorkingFolder' ORDER BY menu_item"
    
    i = i + 1
    arr_controls(i, 0) = "new_nbi_template"
    arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & Load.sources.menu_list_table & " WHERE item_type = 'NBITemplate' ORDER BY menu_item"
    
    i = i + 1
    arr_controls(i, 0) = "new_rp_region_id"
    arr_controls(i, 1) = "SELECT jurisdiction_id id, jurisdiction menu_item FROM " & Load.sources.jurisdictions_view _
    & " WHERE jurisdiction_type = 'rp_region' ORDER BY jurisdiction"
    
    'remove existing values from combo boxes
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    With Forms(str_form)
        For i = 0 To UBound(arr_controls)
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        'add new values to combo boxes
        For i = 0 To UBound(arr_controls)
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
    End With
    
    DoCmd.MoveSize Right:=200, Down:=200, Width:=18800, Height:=12000
    fix_rs.jurisdictions_f
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "open_forms.jurisdictions_f"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
End Sub

Public Sub system_dashboard_f()
    Load.call_stack = Load.call_stack & vbNewLine & "open_forms.system_dashboard_f"
    Dim str_form As String
    global_vars.system_dashboard.init
    str_form = global_vars.system_dashboard.form_name
    
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    
    With Forms(str_form)
        .Caption = "system dashboard"
        .SetFocus
        DoCmd.MoveSize Right:=400, Down:=400, Width:=27000, Height:=16000
    End With
    
    With Forms(str_form)
        Dim control_count As Integer
        Dim arr_controls() As Variant
        ReDim arr_controls(0 To 50, 0 To 1)
        
        control_count = 1
        arr_controls(control_count, 0) = global_vars.system_dashboard.filter_routine_name.field_name
        arr_controls(control_count, 1) = "SELECT DISTINCT(routine_name) menu_item FROM " & Load.sources.log_errors_v & " ORDER BY routine_name"
        
        control_count = control_count + 1
        arr_controls(control_count, 0) = global_vars.system_dashboard.filter_app_name.field_name
        arr_controls(control_count, 1) = "SELECT DISTINCT(app_name) menu_item FROM " & Load.sources.log_errors_v & " ORDER BY app_name"
        
        control_count = control_count + 1
        arr_controls(control_count, 0) = global_vars.system_dashboard.filter_app_continent.field_name
        arr_controls(control_count, 1) = "SELECT DISTINCT(app_continent) menu_item FROM " & Load.sources.log_errors_v & " ORDER BY app_continent"
        
        'remove old values
        Dim i As Integer
        For i = 1 To control_count
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        
        'add empty values
        For i = 1 To control_count
            .Controls(arr_controls(i, 0)).AddItem "_all"
        Next i
        
        'add new values
        Dim str_sql As String
        Dim rs As ADODB.Recordset
        For i = 1 To control_count
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem "'" & rs!menu_item & "'"
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
        
        With global_vars.system_dashboard
            .default_values
            .paint
        End With
    
    End With
    
    fix_rs.system_dashboard_f
    
End Sub
Public Sub awac_excels_f()
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String, arr_controls(0 To 50, 0 To 1), control_count As Integer
    Dim str_sql As String
    
    global_vars.awac_excels.init
    str_form = global_vars.awac_excels.form_name
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    global_vars.awac_excels.paint
    fix_rs.awac_excels_f
    
    With Forms(str_form)
        .Caption = "awac excels"
        .SetFocus
        DoCmd.MoveSize Right:=200, Down:=200, Width:=12000, Height:=8000
        
        control_count = 0
        arr_controls(control_count, 0) = "header_to_field"
        arr_controls(control_count, 1) = Load.sources.menu_lists.yes_no
        
        control_count = control_count + 1
        arr_controls(control_count, 0) = "header_cc_field"
        arr_controls(control_count, 1) = Load.sources.menu_lists.yes_no
        
        control_count = control_count + 1
        arr_controls(control_count, 0) = "header_payment_report"
        arr_controls(control_count, 1) = Load.sources.menu_lists.yes_no
        
        control_count = control_count + 1
        arr_controls(control_count, 0) = "header_v5_report"
        arr_controls(control_count, 1) = Load.sources.menu_lists.yes_no
        
        'remove exusting values
        Dim i As Integer, rs As ADODB.Recordset
        For i = 1 To control_count
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        
        'add new values
        For i = 0 To control_count
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
            rs.Close
        Next i
        
    End With
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "open_forms.awac_excels_f"
        .milestone = "str_sql = " & str_sql
        .params = "control_count = " & control_count & ", i = " & i
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub extended_info_f()
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String
    str_form = "extended_info_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    With Forms(str_form)
        .SetFocus
        DoCmd.MoveSize 200, 200, 24000, 13000
    End With
    
outro:
    Exit Sub
    
err_handler:
    MsgBox Load.system_info.error_instruction & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: open_forms.extended_info_f" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: " & Load.system_info.app_name, , Load.system_info.error_msg_heading
    GoTo outro
End Sub
Public Sub service_messages_f()
    'intro
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String, str_sql As String, rs As ADODB.Recordset, i As Integer
    str_form = "service_messages_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    
    'load recordset for form
    fix_rs.service_messages_f
    
    With Forms(str_form)
        .Caption = "Service messages"
        .SetFocus
        DoCmd.MoveSize Right:=200, Down:=200, Width:=16200, Height:=10000
        
        Dim arr_controls() As Variant, control_count As Integer
        'add items to message type. Done by fetchig unique values from message_type column in dataset
        'redim to reset array
        ReDim arr_controls(0 To 100, 0 To 1)
        i = 0
        arr_controls(i, 0) = "header_message_type": i = i + 1
        arr_controls(i, 0) = "header_message_type_filter": i = i + 1
        
        control_count = i - 1
                   
        'add new items
        str_sql = "SELECT DISTINCT message_type FROM " & Load.sources.service_messages_table & " ORDER BY message_type"
        Set rs = utilities.create_adodb_rs(conn, str_sql)
            For i = 0 To control_count
                'remove old items
                Do While .Controls(arr_controls(i, 0)).ListCount > 0
                    .Controls(arr_controls(i, 0)).RemoveItem (0)
                Loop
                rs.MoveFirst
                Do Until rs.EOF
                    .Controls(arr_controls(i, 0)).AddItem rs!message_type & ";" & rs!message_type
                    rs.MoveNext
                Loop
            Next i
        rs.Close
        
        'add items to message priority. Done by fetchig unique values from message_priority column in dataset
        'redim to reset array
        ReDim arr_controls(0 To 100, 0 To 1)
        i = 0
        arr_controls(i, 0) = "header_message_priority": i = i + 1
        arr_controls(i, 0) = "header_message_priority_filter": i = i + 1
        
        control_count = i - 1
                   
        'add new items
        str_sql = "SELECT DISTINCT message_priority FROM " & Load.sources.service_messages_table & " ORDER BY message_priority"
        Set rs = utilities.create_adodb_rs(conn, str_sql)
            For i = 0 To control_count
                'remove old items
                Do While .Controls(arr_controls(i, 0)).ListCount > 0
                    .Controls(arr_controls(i, 0)).RemoveItem (0)
                Loop
                rs.MoveFirst
                Do Until rs.EOF
                    .Controls(arr_controls(i, 0)).AddItem rs!message_priority & ";" & rs!message_priority
                    rs.MoveNext
                Loop
            Next i
        rs.Close
                    
        'add items to yes/no boxes
        ReDim arr_controls(0 To 100, 0 To 1)
        i = 1
        arr_controls(i, 0) = "header_for_stella_uw_usa_id"
        arr_controls(i, 1) = Load.sources.menu_lists.yes_no
            
        i = i + 1
        arr_controls(i, 0) = "header_for_stella_uw_eur_id"
        arr_controls(i, 1) = Load.sources.menu_lists.yes_no
        
        i = i + 1
        arr_controls(i, 0) = "header_for_stella_admin_eur_id"
        arr_controls(i, 1) = Load.sources.menu_lists.yes_no
        
        i = i + 1
        arr_controls(i, 0) = "header_for_cm_usa_id"
        arr_controls(i, 1) = Load.sources.menu_lists.yes_no
        
        i = i + 1
        arr_controls(i, 0) = "header_for_cm_eur_id"
        arr_controls(i, 1) = Load.sources.menu_lists.yes_no
        
        arr_controls(0, 0) = i
        
        'add new values
        For i = 1 To arr_controls(0, 0)
            'remove old items
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
            
            'add new items
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
            rs.Close
        Next i
        service_messages.reset_input_fields
    End With

outro:
    Exit Sub
    
err_handler:

    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "open_forms.service_messages_f"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub global_policy_list_f()
    Const proc_name As String = "open_forms.global_policy_list_f"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim arr_controls() As Variant
    Dim str_form_name As String
    
    str_form_name = "global_policy_list_f"
    If CurrentProject.AllForms(str_form_name).IsLoaded = False Then
        DoCmd.OpenForm str_form_name
        DoCmd.MoveSize Right:=9250, Down:=0, Width:=17000, Height:=5000
    End If
    
    global_policy_list.init
    
    With Forms(str_form_name)
        .Caption = "Global Policy List"
        'add values to combo boxes
        ReDim arr_controls(0 To 50, 0 To 1)
        
        i = 1
        arr_controls(i, 0) = "comBroker"
        arr_controls(i, 1) = "SELECT id, short_name menu_item FROM broker_firms_v ORDER BY short_name; "
        
        i = i + 1
        arr_controls(i, 0) = "comSector"
        arr_controls(i, 1) = "SELECT sector_id id, sector_name menu_item FROM " & sources.sectors_table _
        & " WHERE sector_type = 494 AND is_deleted = 0 ORDER BY sector_name"
        
        i = i + 1
        arr_controls(i, 0) = "comJurisdiction"
        arr_controls(i, 1) = "SELECT jurisdiction id, jurisdiction menu_item FROM " & sources.jurisdictions_view _
        & " WHERE jurisdiction_type = 'country' ORDER BY jurisdiction_type, jurisdiction"
        
        i = i + 1
        arr_controls(i, 0) = "com_primary_or_xs"
        arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & Load.sources.menu_list_table & " WHERE item_type = 'layer_type'"
        
        i = i + 1
        arr_controls(i, 0) = "com_risk_type"
        arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & Load.sources.menu_list_table & " WHERE item_type = 'RiskType' ORDER BY menu_item"
        
        arr_controls(0, 0) = i
        
        'remove exusting values
        For i = 1 To arr_controls(0, 0)
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        
        'all _all items
        .Controls("com_primary_or_xs").AddItem "-1;_all"
        .Controls("com_risk_type").AddItem "-1;_all"
        .Controls("comSector").AddItem "-1;_all"
        .Controls("comBroker").AddItem "-1;_all"
        !com_sub_sector.AddItem "-1;_all"
        
        'add new values
        For i = 1 To arr_controls(0, 0)
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Dim check_helper As Variant
            check_helper = 666
            Do While rs.EOF = False
                If rs!menu_item <> "_all" Then
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next i
        
    End With
    
    utilities.paint_control str_form_name, global_policy_list.col_navins_homes
    
    fix_rs.global_policy_list_f
    
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
        & "Where: open_forms.global_policy_list_f" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub
Public Sub deal_list_f()
    'intro
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    Load.check_conn_and_variables
    
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    'open form
    Dim str_form As String
    str_form = "deal_list_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
        DoCmd.MoveSize Right:=9250, Down:=0, Width:=19000, Height:=5000
    End If
    
    With Forms(str_form)
        .Caption = "Master Deal List"
        'add values to combo boxes
        Dim arr_controls() As Variant
        ReDim arr_controls(0 To 50, 0 To 1)
        
        i = 1
        arr_controls(i, 0) = "comBroker"
        arr_controls(i, 1) = "SELECT id, short_name menu_item FROM broker_firms_v ORDER BY short_name; "
        
        i = i + 1
        arr_controls(i, 0) = "comSector"
        arr_controls(i, 1) = "SELECT sector_id id, sector_name menu_item FROM " & sources.sectors_table _
        & " WHERE sector_type = 494 AND is_deleted = 0 ORDER BY sector_name"
        
        i = i + 1
        arr_controls(i, 0) = "comJurisdiction"
        arr_controls(i, 1) = "SELECT jurisdiction id, jurisdiction menu_item FROM " & sources.jurisdictions_view _
        & " WHERE jurisdiction_type = 'country' ORDER BY jurisdiction_type, jurisdiction"
        
        i = i + 1
        arr_controls(i, 0) = "com_primary_or_xs"
        arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & Load.sources.menu_list_table & " WHERE item_type = 'layer_type'"
        
        i = i + 1
        arr_controls(i, 0) = "com_risk_type"
        arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & Load.sources.menu_list_table & " WHERE item_type = 'RiskType' ORDER BY menu_item"
        
        arr_controls(0, 0) = i
        
        'remove exusting values
        For i = 1 To arr_controls(0, 0)
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        
        'all all item
        .Controls("com_primary_or_xs").AddItem "-1;_all"
        .Controls("com_risk_type").AddItem "-1;_all"
        .Controls("comSector").AddItem "-1;_all"
        .Controls("comBroker").AddItem "-1;_all"
        !com_sub_sector.AddItem "-1;_all"
        
        'add new values
        For i = 1 To arr_controls(0, 0)
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Do While rs.EOF = False
                If rs!menu_item <> "_all" Then
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next i
        
        .default_filters
    End With
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
        .routine_name = "open_forms.deal_list_f"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
End Sub
Public Sub pricing_guidelines_f()
    'intro
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String
    str_form = "pricing_guidelines_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    
    With Forms(str_form)
        .Caption = "pricing guidelines"
        .SetFocus
        DoCmd.MoveSize Right:=100, Down:=100, Width:=9400, Height:=8000
        .FormFooter.Height = 100
        'remove any existing lists
        Dim i As Integer
        Dim arr_controls() As Variant
        ReDim arr_controls(0 To 50, 0 To 1)
        i = 1
        arr_controls(i, 0) = "header_jurisdiction_id"
        arr_controls(i, 1) = "SELECT jurisdiction_id id, jurisdiction menu_item FROM " & sources.jurisdictions_table _
            & " WHERE is_deleted = 0 AND jurisdiction_type = 'rp_region' ORDER BY jurisdiction": i = i + 1
        arr_controls(i, 0) = "header_super_sector_id"
        arr_controls(i, 1) = "SELECT sector_id id, sector_name menu_item FROM " & Load.sources.sectors_table _
            & " WHERE is_deleted = 0 AND sector_type = 494 ORDER BY sector_name": i = i + 1
        
        arr_controls(0, 0) = i - 1
        For i = 1 To arr_controls(0, 0)
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        'add new values
        Dim str_sql As String
        Dim rs As ADODB.Recordset
        For i = 1 To arr_controls(0, 0)
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Do While rs.EOF = False
                If rs!menu_item <> "_All" Then
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & "'" & rs!menu_item & "'"
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
    End With
    
    'fix rs for form
    fix_rs.pricing_guidelines_f
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: open_forms.pricing_guidelines_f" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
End Sub
Public Sub working_on_it_f(ByVal input_text As String, Optional ByVal text_details As String, Optional ByVal form_height As Long)
    Load.call_stack = Load.call_stack & vbNewLine & "open_forms.working_on_it_f"
    Dim str_form As String
    On Error Resume Next
    str_form = Load.form_names.working_on_it_f
    
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
    Else
        'working on it already showing, so do nothing
        GoTo outro
    End If
    
    With Forms(str_form)
        !input_text.Value = input_text
        !text_details = text_details
        !placeholder.SetFocus
    End With
    Dim int_height As Long
    int_height = 4000
    If form_height > 0 Then int_height = form_height
    DoCmd.MoveSize Right:=200, Down:=200, Width:=10000, Height:=int_height
    Forms(str_form).Repaint
outro:
    Exit Sub
    
End Sub
Public Sub working_on_it_f__close()
    On Error Resume Next
    If CurrentProject.AllForms(Load.form_names.working_on_it_f).IsLoaded = True Then DoCmd.Close acForm, Load.form_names.working_on_it_f
End Sub
Public Sub broker_persons_f()
    Load.call_stack = Load.call_stack & vbNewLine & "open_forms.broker_persons_f"
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String
    str_form = "broker_person_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    fix_rs.broker_persons_f
    
    'add values to combo boxes
    With Forms(str_form)
        Dim control_count As Integer
        Dim arr_controls() As Variant
        ReDim arr_controls(0 To 50, 0 To 1)
        
        control_count = 1
        arr_controls(control_count, 0) = "new_broker_firm_id"
        arr_controls(control_count, 1) = "SELECT id, short_name menu_item FROM " & sources.broker_firms_view & " WHERE NOT broker_firm_id = 12 ORDER BY short_name"
        
        arr_controls(0, 0) = control_count
        
        'remove old values
        Dim i As Integer
        For i = 1 To arr_controls(0, 0)
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        
        'add new values
        Dim str_sql As String
        Dim rs As ADODB.Recordset
        For i = 1 To arr_controls(0, 0)
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
    End With

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
        .call_stack = Load.call_stack
        .routine_name = "open_forms.broker_persons_f"
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
    Dim proc_name As String
    proc_name = "open_forms.broker_firms_f"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String
    str_form = "broker_firms_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    fix_rs.broker_firms_f
    
    'add values to combo boxes
    With Forms(str_form)
        .SetFocus
        DoCmd.MoveSize Right:=300, Down:=300, Width:=13700, Height:=10000
        !new_broker_firm_id = ""
        Dim control_count As Integer
        Dim arr_controls() As Variant
        ReDim arr_controls(0 To 50, 0 To 1)
        
        control_count = 1
        arr_controls(control_count, 0) = "new_jurisdiction_id"
        arr_controls(control_count, 1) = Load.sources.menu_lists.country_list
        
        'remove old values
        Dim i As Integer
        For i = 1 To control_count
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        
        'add new values
        Dim str_sql As String
        Dim rs As ADODB.Recordset
        For i = 1 To control_count
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Do While rs.EOF = False
                If rs!menu_item <> "_all" Then
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
    End With

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
        .routine_name = "open_forms.broker_firms_f"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
    GoTo outro
End Sub

Public Sub brands_f()
    Const proc_name As String = "open_forms.brands_f"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If is_debugging = True Then On Error GoTo 0
    
    Dim objFormatConds As FormatCondition
    Dim str_form As String
    
    str_form = "brands_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    
    With Forms(str_form)
        .Caption = "manage brands"
        .SetFocus
        DoCmd.MoveSize Right:=200, Down:=200, Width:=15000, Height:=8000
        .FormFooter.Height = 100
        
        Dim arr_controls() As Variant, i As Integer
        ReDim arr_controls(0 To 50, 0 To 1)
        i = 1
        arr_controls(i, 0) = "header_is_active"
        arr_controls(i, 1) = Load.sources.menu_lists.yes_no
        
        arr_controls(0, 0) = i
        For i = 1 To arr_controls(0, 0)
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        'add new values
        On Error GoTo err_handler
        Dim str_sql As String
        Dim rs As ADODB.Recordset
        For i = 1 To arr_controls(0, 0)
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
        
        With .Controls("is_advisor")
            .FormatConditions.Delete
            
            Set objFormatConds = .FormatConditions.Add(acExpression, , "[is_advisor] = 1")
            .FormatConditions(0).BackColor = Load.colors.light_green
        End With
        
        With .Controls("is_broker")
            .FormatConditions.Delete
            
            Set objFormatConds = .FormatConditions.Add(acExpression, , "[is_broker] = 1")
            .FormatConditions(0).BackColor = Load.colors.light_green
        End With
        
        With .Controls("is_competitor")
            .FormatConditions.Delete
            
            Set objFormatConds = .FormatConditions.Add(acExpression, , "[is_competitor] = 1")
            .FormatConditions(0).BackColor = Load.colors.light_green
        End With
        
        With .Controls("is_carrier")
            .FormatConditions.Delete
            
            Set objFormatConds = .FormatConditions.Add(acExpression, , "[is_carrier] = 1")
            .FormatConditions(0).BackColor = Load.colors.light_green
        End With
        
        With .Controls("is_party")
            .FormatConditions.Delete
            
            Set objFormatConds = .FormatConditions.Add(acExpression, , "[is_party] = 1")
            .FormatConditions(0).BackColor = Load.colors.light_green
        End With

    End With
    
    fix_rs.brands_f True
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: open_forms.brands_f" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"

End Sub
Public Sub uws_role_f()
    Const proc_name As String = "open_forms.uws_role_f"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    user_management.init
    
    Dim arr_controls() As Variant
    Dim i As Long
    Dim rs As ADODB.Recordset
    Dim str_form As String
    Dim str_sql As String
        
    str_form = "uws_role_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
        DoCmd.MoveSize Right:=200, Down:=200, Width:=7000, Height:=5000
    End If
    
    With Forms(str_form)
        
        ReDim arr_controls(0 To 100, 0 To 1)
        i = 0
        i = i + 1
        arr_controls(i, 0) = "employee_role_id__employee_roles_t"
        arr_controls(i, 1) = Load.sources.menu_lists.employee_roles
        
        arr_controls(0, 0) = i
        
        For i = 1 To arr_controls(0, 0)
            .Controls(arr_controls(i, 0)).RowSource = ""
        Next i
        
        'add new values
        For i = 1 To arr_controls(0, 0)
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
    End With
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub

Public Sub uws_add_f()
    Const proc_name As String = "open_forms.uws_add_f"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    user_management.init
    
    Dim arr_controls() As Variant
    Dim i As Long
    Dim rs As ADODB.Recordset
    Dim str_form As String
    Dim str_sql As String
        
    str_form = "uws_add_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
        DoCmd.MoveSize Right:=200, Down:=200, Width:=15000, Height:=10000
    End If
    
    With Forms(str_form)
        .Controls(user_management.uw_data_controls.budget_home_id.field_name).RowSource = Load.sources.menu_lists.row_source_budget_homes
        
        ReDim arr_controls(0 To 100, 0 To 1)
        
        i = 1
        arr_controls(i, 0) = "usertype_id"
        arr_controls(i, 1) = "SELECT menu_id id, menu_item FROM " & Load.sources.menu_list_table & " WHERE item_type = 'UserType'"
        
        i = i + 1
        arr_controls(i, 0) = "employee_role_id"
        arr_controls(i, 1) = Load.sources.menu_lists.employee_roles
        
        arr_controls(0, 0) = i
        
        For i = 1 To arr_controls(0, 0)
            .Controls(arr_controls(i, 0)).RowSource = ""
        Next i
        
        'add new values
        For i = 1 To arr_controls(0, 0)
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
    End With
    
    'set default values
    With user_management.uw_data_controls
        Forms(str_form).Controls(.is_employed_id.field_name) = 93
        Forms(str_form).Controls(.can_change_budget_home_id.field_name) = 94
        Forms(str_form).Controls(.can_change_general_id.field_name) = 94
        Forms(str_form).Controls(.can_change_jurisdictions_id.field_name) = 94
        Forms(str_form).Controls(.can_change_uws_id.field_name) = 94
        Forms(str_form).Controls(.has_admin_access_id.field_name) = 94
        Forms(str_form).Controls(.is_dev_id.field_name) = 94
        Forms(str_form).Controls(.is_regional_lead_id.field_name) = 94
    End With
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub uws_f()
    Const proc_name As String = "open_forms.uws_f"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    user_management.init
    
    Dim rs As ADODB.Recordset
    Dim str_form As String
    Dim str_sql As String
        
    str_form = "uws_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then
        DoCmd.OpenForm str_form
        DoCmd.MoveSize Right:=200, Down:=200, Width:=16500, Height:=16000
    End If
    
    With Forms(str_form)
        .Controls(user_management.header_controls.header_chooser_year.field_name).Value = 0
    End With
    'fix rs for form
    fix_rs.uws_f " AND (user_type_id = 150 OR user_type_id = 149)", Load.sources.uw_roles_view
    user_management.uw_statistics " AND (user_type_id = 150 OR user_type_id = 149)"
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "open_forms.uws_f"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With

End Sub
Public Sub lawyers_f()
    On Error GoTo err_handler
    
    Dim arr_controls() As Variant
    Dim control_count As Integer
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    Dim str_form As String
    
    str_form = "lawyers_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    fix_rs.lawyers_f
    
    'add values to combo boxes
    With Forms(str_form)
        !new_lawyer_id = ""
        
        ReDim arr_controls(0 To 50, 0 To 1)
        
        control_count = 1
        arr_controls(control_count, 0) = "new_law_firm"
        arr_controls(control_count, 1) = "SELECT id, firm_name menu_item FROM " & sources.law_firms_view & " ORDER BY firm_name"
        
        control_count = control_count + 1
        arr_controls(control_count, 0) = "new_jurisdiction"
        arr_controls(control_count, 1) = "SELECT jurisdiction_id id, jurisdiction menu_item FROM " & sources.jurisdictions_view & " ORDER BY jurisdiction"
        
        control_count = control_count + 1
        arr_controls(control_count, 0) = "new_is_counsel"
        arr_controls(control_count, 1) = Load.sources.menu_lists.yes_no
        
        'remove old values
        
        For i = 1 To control_count
            .Controls(arr_controls(i, 0)).RowSource = ""
        Next i
        
        'add new values
        
        For i = 1 To control_count
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
                Do While rs.EOF = False
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";'" & rs!menu_item & "'"
                    rs.MoveNext
                Loop
            rs.Close
        Next i
        Set rs = Nothing
    End With

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
        & "Where: open_forms.broker_persons_f" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub
Public Sub financial_advisors_f()
    'intro
    On Error GoTo err_handler
    
    Dim str_form As String: str_form = "add_financial_advisor_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    fix_rs.financial_advisors_f
    With Forms(str_form)
        !new_id = ""
        Dim control_count As Integer
        Dim arr_controls() As Variant
        ReDim arr_controls(0 To 50, 0 To 1)
        control_count = 0
        arr_controls(control_count, 0) = "new_country"
        arr_controls(control_count, 1) = "SELECT jurisdiction_id id, jurisdiction menu_item FROM " & sources.jurisdictions_view & " ORDER BY jurisdiction"
        control_count = control_count + 1
        
        'remove old values
        Dim i As Integer
        For i = 0 To control_count - 1
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        
        'add new values
        Dim str_sql As String
        Dim rs As ADODB.Recordset
        For i = 0 To control_count - 1
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Do While rs.EOF = False
                .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
    End With

outro:
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: open_forms.financial_advisors_f" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub
Public Sub law_firms_f()
    Dim proc_name As String
    proc_name = "open_forms.law_firms_f"
    Load.call_stack = proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.check_conn_and_variables
    
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim str_form As String
    Dim str_sql As String
    
    str_form = "law_firms_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    fix_rs.law_firms_f
    With Forms(str_form)
        !new_id = ""
        Dim control_count As Integer
        Dim arr_controls() As Variant
        ReDim arr_controls(0 To 50, 0 To 1)
        
        control_count = 1
        arr_controls(control_count, 0) = "header_jurisdiction"
        arr_controls(control_count, 1) = "SELECT jurisdiction_id id, jurisdiction menu_item FROM " _
        & sources.jurisdictions_view & " ORDER BY jurisdiction"
        
        control_count = control_count + 1
        arr_controls(control_count, 0) = "header_is_counsel"
        arr_controls(control_count, 1) = Load.sources.menu_lists.yes_no
        
        'remove old values
        For i = 1 To control_count
            .Controls(arr_controls(i, 0)).RowSource = ""
        Next i
        
        'add new values
        For i = 1 To control_count
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Dim check_helper As Variant
            check_helper = -1
            Do While rs.EOF = False
                'several items don't have a menu info, and vba cannot check if fields exists (nor can it try-catch)
                On Error Resume Next
                    check_helper = rs!menu_info
                On Error GoTo err_handler
                If check_helper = -1 Then
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                Else
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item & ";'" & rs!menu_info & "'"
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
    End With
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
        & "Where: open_forms.law_firms_f" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub
Public Sub templates_f()
    On Error GoTo err_handler
    If Load.is_init = False Then Central.stella_init
    
    Dim str_form As String: str_form = "add_templates_f"
    If CurrentProject.AllForms(str_form).IsLoaded = False Then DoCmd.OpenForm str_form
    fix_rs.templates_f
    With Forms(str_form)
        .SetFocus
        DoCmd.MoveSize Right:=200, Down:=200, Width:=16000, Height:=6000
        .FormFooter.Height = 100
        !new_id = ""
        Dim control_count As Integer
        Dim arr_controls() As Variant
        ReDim arr_controls(0 To 50, 0 To 1)
        control_count = 0
        arr_controls(control_count, 0) = "new_file_name"
        arr_controls(control_count, 1) = "SELECT id, file_name menu_item FROM " & sources.templates_table & " ORDER BY file_name"
        control_count = control_count + 1
        
        'remove old values
        Dim i As Integer
        For i = 0 To control_count - 1
            Do While .Controls(arr_controls(i, 0)).ListCount > 0
                .Controls(arr_controls(i, 0)).RemoveItem (0)
            Loop
        Next i
        
        'add items to folder combo box
        Dim template_folder As Object, folder As Object, templates_folder_path As String
        templates_folder_path = Load.system_info.system_paths.template_path & "stella_templates\"
        With !new_folder
            Do While .ListCount > 0
                .RemoveItem (0)
            Loop
            Set template_folder = CreateObject("Scripting.FileSystemObject")
            For Each folder In template_folder.getfolder(templates_folder_path).SubFolders
                .AddItem CStr(folder.name)
            Next folder
            Set template_folder = Nothing
        End With
        
        'add new values
        Dim str_sql As String
        Dim rs As ADODB.Recordset
        For i = 0 To control_count - 1
            str_sql = arr_controls(i, 1)
            Set rs = utilities.create_adodb_rs(conn, str_sql)
            Dim check_helper As Variant
            check_helper = -1
            Do While rs.EOF = False
                'several items don't have a menu info, and vba cannot check if fields exists (nor can it try-catch)
                If check_helper = -1 Then
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item
                Else
                    .Controls(arr_controls(i, 0)).AddItem rs!id & ";" & rs!menu_item & ";'" & rs!menu_info & "'"
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next i
        Set rs = Nothing
    End With

outro:
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: open_forms.templates_f" & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
    
End Sub