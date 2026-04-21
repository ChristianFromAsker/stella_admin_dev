Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public awac_conversion_path As String
Public base_path As String
Public base_path_eur As String
Public base_path_us As String
Public bound_folder As String

Public cm_admin As String
Public common_path As String
Public common_path_eur As String
Public common_path_us As String

Public images_path As String
Public individual_path_eur As String
Public individual_path_us As String
Public intranet As String

Public logs_folder As String
Public log_for_folder_moving As String
Public messages_path As String
Public non_bound_folder As String

Public old_stella_versions As String
Public placeholder_path_published_eur As String
Public placeholder_path_published_us As String
Public published As String
Public published_eur As String
Public published_us As String
Public pws As String

Public rating_tool_template As String
Public reports_app As String
Public root_deal_path As String

Public scripts_path As String
Public stable_builds As String
Public stable_builds_us As String
Public stella_app_location_general As String
Public stella_app_location_general_network As String
Public stella_app_location_individual As String
Public stella_app_location_individual_network As String
Public stella_exports As String
Public stella_path As String
Public stella_path_eur As String
Public stella_path_us As String
Public stella_templates As String

Public template_path As String

Public working_folder As String
Public Sub init()
    Const proc_name As String = "cls_system_paths.init"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim drive_path As String
    Dim drive_path_eur As String
    Dim drive_path_us As String
    Dim str_milestone As String
    
    With Me
        'International
        str_milestone = "1"
        drive_path_eur = "\\file01-weu-prd\data$\Data\"
        .base_path_eur = drive_path_eur & "R i s k P o i n t\Denmark\35 M&A\"
        
        .awac_conversion_path = .stella_path_eur & "awac_excel\"
        .bound_folder = base_path_eur & "M&A - Bound\"
        .non_bound_folder = base_path_eur & "M&A - Non-bound\"
        .working_folder = .base_path_eur & "M&A - Working\"
        
        .stella_path_eur = .base_path_eur & "Stella\"
        .common_path_eur = .stella_path_eur & "common\"
        .individual_path_eur = .stella_path_eur & "\published\individual\"
        .old_stella_versions = .stella_path_eur & "old_versions\"
        .published_eur = .stella_path_eur & "published\"
        .placeholder_path_published_eur = .published_eur & "placeholders\"
        
        'Americas
        str_milestone = "2"
        drive_path_us = "\\C325-us-file01\325-us-data$\Data\"
        .base_path_us = drive_path_us & "system\"
        
        .stella_path_us = .base_path_us & "Stella\"
        .common_path_us = .stella_path_us & "common\"
        .individual_path_us = .stella_path_us & "published\individual\"
        .published_us = .stella_path_us & "published\"
        .placeholder_path_published_us = .published_us & "placeholders\"
        
        'logical
        str_milestone = "3"
        If Load.system_info.ip_current = Load.system_info.continents.eurasia_ip Then
            .base_path = .base_path_eur
            root_deal_path = drive_path & .base_path_eur
            template_path = base_path & "Templates\"
            .stella_app_location_individual = .individual_path_eur
        ElseIf Load.system_info.ip_current = Load.system_info.continents.america_ip Then
            .base_path = .base_path_us
            .root_deal_path = drive_path_us & "deals\"
            .template_path = base_path_us & "stella_templates\"
            .stella_app_location_individual = .individual_path_us
        Else
            '1 September 2025, CK: This should only happen when CK test Stella outside the AVD.
            drive_path = "F:\"
            base_path = drive_path & "R i s k P o i n t\Denmark\35 M&A\"
            root_deal_path = drive_path & "R i s k P o i n t\Denmark\35 M&A\"
            template_path = base_path & "Templates\"
        End If
        
        str_milestone = "4"
        .stella_path = .base_path & "Stella\"
        .common_path = .stella_path & "common\"
        .images_path = .stella_path & "images\stella_admin\"
        .messages_path = .stella_path & "messages\"
        .published = .stella_path & "published\"
        .rating_tool_template = template_path & "rating_tool\rating_tool.xlsm"
        
        .pws = .stella_path & "pws"
        .stable_builds = .stella_path & "stable_builds\"
        .stable_builds_us = .base_path_us & "stella\stable_builds\"
        .stella_exports = .stella_path & "stella_exports\"
        .stella_templates = .template_path & "stella_templates\"
    
        'file paths
        .reports_app = .common_path & "reports.accdb"
        .scripts_path = .common_path & "scripts_01.accdb"
        .cm_admin = .published_eur & "cm_admin.accdb"
        .logs_folder = .stella_path & "logs\"
        .log_for_folder_moving = .logs_folder & "log for folder moving.txt"
    
    End With
    
outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:

    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: " & proc_name & vbNewLine _
        & "Milestone: " & str_milestone & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "Call stack: " & Load.call_stack & vbNewLine _
        & "App: Stella Admin", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
    
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = "cls_system_paths.init"
        .milestone = str_milestone
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = False
        .send_error err_object
    End With
End Sub