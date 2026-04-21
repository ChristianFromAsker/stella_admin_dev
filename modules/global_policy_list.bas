Option Compare Database
Option Explicit

Public col_navins_homes As New Collection
Public navins_canada As New cls_field
Public navins_denmark As New cls_field
Public navins_dubai As New cls_field
Public navins_finland As New cls_field
Public navins_germany As New cls_field
Public navins_holland As New cls_field
Public navins_norway As New cls_field
Public navins_singapore As New cls_field
Public navins_spain As New cls_field
Public navins_sweden As New cls_field
Public navins_switzerland As New cls_field
Public navins_uk_old As New cls_field
Public navins_uk_solutions As New cls_field
Public navins_usa As New cls_field

Public is_init As Boolean
Public sort_field As String
Public sort_order As String

Public Sub update_navins_home_button(ByRef rp_entity As cls_field)
    Const proc_name As String = "global_policy_list.navins_home_button"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    With rp_entity
        If .field_value = True Then
            .field_value = False
            .font_color = Load.colors.black
            .field_bg_color = Load.colors.inactive_navins_home
            .field_caption = Split(.field_caption, ":")(0) & ": off"
        Else
            .field_value = True
            .font_color = Load.colors.white
            .field_bg_color = Load.colors.active_navins_home
            .field_caption = Split(.field_caption, ":")(0) & ": on"
        End If
    End With
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Function date_control() As Boolean
    'intro
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String
    str_form = "global_policy_list_f"
    
    With Forms(str_form)
        If CInt(!date_year_end) < CInt(!date_year_start) Then
            date_control = False
            !date_year_end.BackColor = Load.colors.yellow
            GoTo outro
        End If
        If CInt(!date_year_end) = CInt(!date_year_start) And CInt(!date_month_end) < CInt(!date_month_start) Then
            date_control = False
            !date_month_end.BackColor = Load.colors.yellow
            GoTo outro
        End If
        
        If !date_month_start > 12 Then
            date_control = False
            !date_month_start.BackColor = Load.colors.yellow
            GoTo outro
        End If
        
        If !date_month_end > 12 Then
            date_control = False
            !date_month_end.BackColor = Load.colors.yellow
            GoTo outro
        End If
        
        !date_month_end.BackColor = Load.colors.white
        !date_year_end.BackColor = Load.colors.white
        date_control = True
    End With
outro:
    Exit Function
err_handler:
    MsgBox "Something went wrong. Snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: global_policy_list.date_control" & vbNewLine _
        & "Parameters: n/a " & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Function
Public Sub set_default_values()
    Const proc_name As String = "global_policy_list.set_default_values"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_form As String
    str_form = "global_policy_list_f"
    With Forms(str_form)
        !comJurisdiction = "_All"
        !com_risk_type = -1
        !comSector = -1
        !com_sub_sector = -1
        !com_primary_or_xs = -1
        !date_year_start = Year(Date)
        !date_year_end = Year(Date)
        !date_month_start = 1
        !date_month_end = 12
        !deal_name_search = ""
    End With
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    MsgBox "Something went wrong. Snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: global_policy_list.set_default_values" & vbNewLine _
        & "Parameters: n/a " & vbNewLine _
        & "App: Stella Admin Eur", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    Resume outro
End Sub
Public Sub init()
    is_init = True
    sort_field = "inception_date"
    sort_order = "ASC"
    
    With global_policy_list.navins_canada
        .field_name = "cmd_canada"
        .id = 435
    End With
    With global_policy_list.navins_denmark
        .field_name = "cmd_denmark"
        .id = 76
    End With
    With global_policy_list.navins_dubai
        .field_name = "cmd_dubai"
        .id = 436
    End With
    With global_policy_list.navins_finland
        .field_name = "cmd_finland"
        .id = 82
    End With
    With global_policy_list.navins_germany
        .field_name = "cmd_germany"
        .id = 79
    End With
    With navins_holland
        .field_name = "cmd_holland"
        .id = 83
    End With
    With navins_norway
        .field_name = "cmd_norway"
        .id = 77
    End With
    With navins_spain
        .field_name = "cmd_spain"
        .id = 87
    End With
    With navins_singapore
        .field_name = "cmd_singapore"
        .id = 86
    End With
    With navins_sweden
        .field_name = "cmd_sweden"
        .id = 78
    End With
    With navins_switzerland
        .field_name = "cmd_switzerland"
        .id = 95
    End With
    With navins_uk_solutions
        .field_name = "cmd_uk_solutions"
        .id = 437
    End With
    With navins_uk_old
        .field_name = "cmd_uk_old"
        .id = 81
    End With
    With navins_usa
        .field_name = "cmd_usa"
        .id = 93
    End With
    
    Set global_policy_list.col_navins_homes = Nothing
    With global_policy_list.col_navins_homes
        .Add global_policy_list.navins_canada
        .Add global_policy_list.navins_denmark
        .Add global_policy_list.navins_dubai
        .Add global_policy_list.navins_finland
        .Add global_policy_list.navins_germany
        .Add global_policy_list.navins_holland
        .Add global_policy_list.navins_norway
        .Add global_policy_list.navins_singapore
        .Add global_policy_list.navins_spain
        .Add global_policy_list.navins_sweden
        .Add global_policy_list.navins_switzerland
        .Add global_policy_list.navins_uk_old
        .Add global_policy_list.navins_uk_solutions
        .Add global_policy_list.navins_usa
    End With
    
    Dim field_form As cls_field
    Dim field_form_2 As cls_field
    
    For Each field_form In global_policy_list.col_navins_homes
        With field_form
            .field_bg_color = Load.colors.active_navins_home
            .font_color = Load.colors.white
            .field_value = True
            .field_visible = True
            .visible_in_europe = True
            .visible_in_america = True
        End With
    Next field_form
    
    For Each field_form In global_vars.col_rp_entities
        For Each field_form_2 In global_policy_list.col_navins_homes
            If field_form_2.id = field_form.id Then
                field_form_2.field_caption = Right(field_form.field_caption, Len(field_form.field_caption) - 3) & ": on"
            End If
        Next field_form_2
    Next field_form
outro:
    Exit Sub
End Sub
    