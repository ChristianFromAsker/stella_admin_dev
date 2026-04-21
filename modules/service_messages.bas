Option Compare Database
Option Explicit

Public Sub reset_input_fields()
    Dim input_controls() As Variant, str_form As String, i As Integer
    str_form = "service_messages_f"
    
    input_controls = service_messages.declare_fields
    With Forms(str_form)
        For i = 1 To input_controls(0, 0)
            .Controls("header_" & input_controls(i, 0)) = ""
            If input_controls(i, 2) = menu_list.yes Then
                .Controls("header_" & input_controls(i, 0)).BackColor = colors.white
            End If
        Next i
        !cmd_extended_info.Visible = False
        !header_extended_info.Visible = False
        !header_extended_info_lbl.Visible = False
        !cmd_add_change.Caption = "add"
    End With
End Sub

Public Function declare_fields()
    
    Dim i As Integer
    Dim input_controls(0 To 100, 0 To 2) As Variant
    Const is_mandatory As Integer = 1
    Const change_bg_color As Integer = 2
    ' the second value indicates if value is mandatory (1) or not (0)
    i = 1
    input_controls(i, 0) = "message_id"
    input_controls(i, is_mandatory) = menu_list.no
    input_controls(i, change_bg_color) = menu_list.no
    i = i + 1
    
    input_controls(i, 0) = "message_text"
    input_controls(i, is_mandatory) = menu_list.yes
    input_controls(i, change_bg_color) = menu_list.yes
    i = i + 1
    
    input_controls(i, 0) = "message_type"
    input_controls(i, is_mandatory) = menu_list.yes
    input_controls(i, change_bg_color) = menu_list.yes
    i = i + 1
    
    input_controls(i, 0) = "message_priority"
    input_controls(i, is_mandatory) = menu_list.yes
    input_controls(i, change_bg_color) = menu_list.yes
    i = i + 1
    
    input_controls(i, 0) = "for_stella_uw_eur_id"
    input_controls(i, is_mandatory) = menu_list.yes
    input_controls(i, change_bg_color) = menu_list.yes
    i = i + 1
    
    input_controls(i, 0) = "for_stella_uw_usa_id"
    input_controls(i, is_mandatory) = menu_list.yes
    input_controls(i, change_bg_color) = menu_list.yes
    i = i + 1
    
    input_controls(i, 0) = "for_cm_eur_id"
    input_controls(i, is_mandatory) = menu_list.yes
    input_controls(i, change_bg_color) = menu_list.yes
    i = i + 1
    
    input_controls(i, 0) = "for_cm_usa_id"
    input_controls(i, is_mandatory) = menu_list.yes
    input_controls(i, change_bg_color) = menu_list.yes
    i = i + 1
    
    input_controls(i, 0) = "for_stella_admin_eur_id"
    input_controls(i, is_mandatory) = menu_list.yes
    input_controls(i, change_bg_color) = menu_list.yes
    i = i + 1
    
    input_controls(i, 0) = "start_date"
    input_controls(i, is_mandatory) = menu_list.yes
    input_controls(i, change_bg_color) = menu_list.yes
    i = i + 1
    
    input_controls(i, 0) = "end_date"
    input_controls(i, is_mandatory) = menu_list.yes
    input_controls(i, change_bg_color) = menu_list.yes
    i = i + 1
    
    input_controls(0, 0) = i - 1
    declare_fields = input_controls
End Function