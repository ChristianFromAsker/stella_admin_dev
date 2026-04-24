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
Public cm_path As String
Public cm_path_file As String
Public common_path As String
Public common_path_eur As String
Public common_path_us As String

Public guides As String
Public images_path As String
Public images_path_stella_admin As String
Public individual_path_eur As String
Public individual_path_us As String

Public loader_path As String
Public log_call_stack_file As String
Public log_for_folder_moving As String
Public logs_folder As String

Public mysql_backup_path As String
Public messages_path As String
Public non_bound_folder As String

Public old_stella_versions As String

Public placeholder_path As String
Public placeholder_path_published_eur As String
Public placeholder_path_published_us As String
Public published_docs As String
Public published_eur As String
Public published_us As String
Public pws As String

Public rating_tool_template As String
Public reports_app As String
Public root_deal_path As String

Public scripts_path As String
Public scripts_path_secondary As String
Public snake_game As String
Public stable_builds As String
Public stable_builds_us As String
Public stella_app_location_general As String
Public stella_app_location_general_network As String
Public stella_app_location_individual As String
Public stella_app_location_individual_network As String
Public stella_exports As String
Public stella_images_cm As String
Public stella_path As String
Public stella_path_eur As String
Public stella_path_us As String
Public stella_templates As String

Public template_path As String
Public temp_folder As String

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
        'international
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
        
        'americas
        drive_path_us = "\\C325-us-file01\325-us-data$\Data\"
        .base_path_us = drive_path_us & "system\"
        
        .stella_path_us = .base_path_us & "Stella\"
        .common_path_us = .stella_path_us & "common\"
        .individual_path_us = .stella_path_us & "published\individual\"
        .published_us = .stella_path_us & "published\"
        .placeholder_path_published_us = .published_us & "placeholders\"
        
        'logical
        If Load.system_info.ip_current = Load.system_info.continents.eurasia_ip Then
            drive_path = "\\file01-weu-prd\data$\Data\"
            base_path = drive_path & "R i s k P o i n t\Denmark\35 M&A\"
            root_deal_path = drive_path & "R i s k P o i n t\Denmark\35 M&A\"
            template_path = base_path & "Templates\"
            loader_path = .base_path & "stella\published\Stella UW.accdb"
        ElseIf Load.system_info.ip_current = Load.system_info.continents.america_ip Then
            drive_path = "\\C325-us-file01\325-us-data$\Data\"
            base_path = drive_path & "system\"
            root_deal_path = drive_path & "deals\"
            template_path = base_path & "stella_templates\"
            loader_path = base_path & "stella_us"
        Else
            '1 September 2025, CK: This should only happen when CK test Stella outside the AVD.
            drive_path = "F:\"
            base_path = drive_path & "R i s k P o i n t\Denmark\35 M&A\"
            root_deal_path = drive_path & "R i s k P o i n t\Denmark\35 M&A\"
            template_path = base_path & "Templates\"
            loader_path = .base_path & "stella\published\Stella UW.accdb"
        End If
        
        .bound_folder = base_path & "M&A - Bound\"
        .non_bound_folder = base_path & "M&A - Non-bound\"
        
        .working_folder = base_path & "M&A - Working\"
        .guides = base_path & "guides\"
        
        .stella_path = base_path & "Stella\"
        .awac_conversion_path = stella_path & "awac_excel\"
        .cm_path = stella_path & "common\"
        .common_path = .stella_path & "common\"
        .images_path = .stella_path & "images\"
        .images_path_stella_admin = .images_path & "stella_admin\"
        .log_call_stack_file = .stella_path & "logs\call_stack.txt"
        .logs_folder = .stella_path & "logs\"
        .log_for_folder_moving = .logs_folder & "log for folder moving.txt"
        .messages_path = .stella_path & "messages\"
        .mysql_backup_path = .stella_path & "mysql_backups\"
        .placeholder_path = .stella_path & "published\placeholders\"
        .published_docs = .stella_path & "published\docs\"
        .pws = .stella_path & "pws"
        .stable_builds = .stella_path & "stable_builds\"
        .stella_app_location_general = .stella_path & "\published"
        .stella_app_location_individual = .stella_path & "\published\individual\"
        .stella_images_cm = .stella_path & "images\cm\"
        .temp_folder = .stella_path & "temp\"
        
        'file paths
        .cm_admin = .published_eur & "cm_admin.accdb"
        .cm_path_file = cm_path & "cm_uw.accdb"
        .rating_tool_template = .template_path & "rating_tool\rating_tool.xlsm"
        .reports_app = .common_path & "reports.accdb"
        .scripts_path = .common_path & "scripts.accdb"
        .scripts_path_secondary = .common_path & "scripts_01.accdb"
        .snake_game = .stable_builds & "snake.xlsm"
        
    End With
outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub