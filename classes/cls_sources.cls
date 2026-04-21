Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public approval_rights_view As String
Public awac_excels_table As String
Public awac_excels_view As String
Public brands_table As String
Public brands_view As String
Public broker_firms_table As String
Public broker_firms_view As String
Public broker_persons_table As String
Public broker_persons_view As String
Public currencies_table As String
Public data_log_table As String
Public deal_roles_table As String
Public deals_for_auto_archiving_view As String
Public deals_table As String
Public deals_table_id As Long
Public deals_view As String

Public employee_role_links_table As String
Public employee_roles_table As String
Public entities_table As String

Public financial_advisors_table As String
Public financial_advisors_view As String
Public global_deal_list_view As String
Public global_policy_list_view As String
Public insurers_table As String
Public insurers_view As String
Public jurisdictions_table As String
Public jurisdictions_view As String
Public law_firms_table As String
Public law_firms_view As String
Public lawyers_table As String
Public lawyers_view As String
Public log_activity_table As String
Public log_errors_eur_t As String
Public log_errors_us_t As String
Public log_errors_v As String
Public menu_list_table As String
Public pricing_guidelines_table As String
Public pricing_guidelines_view As String
Public sectors_table As String
Public service_messages_table As String
Public service_messages_view As String
Public templates_table As String
Public uws_table As String
Public uw_roles_view As String
Public uws_view As String
Public uw_role_fte_by_year_view As String
Public menu_lists As New cls_sources__menu_lists

Public Sub init()
    With Me
        
        .approval_rights_view = "stella_us.approval_rights_v"
    
        .awac_excels_table = "stella_common.awac_excels_t"
        .awac_excels_view = "stella_common.awac_excels_v"
        .broker_persons_view = "broker_persons_v"
        .broker_persons_table = "broker_persons_t"
        .brands_table = "stella_common.brands_t"
        .brands_view = "stella_common.brands_v"
        .broker_firms_table = "broker_firms_t"
        .broker_firms_view = "broker_firms_v"

        .currencies_table = "stella_common.currencies_t"
        .data_log_table = "data_log_t"
        .deal_roles_table = "stella_common.deal_roles_t"
        .deals_for_auto_archiving_view = "deals_for_auto_archiving_v"
        .deals_table = "deals_t"
        .deals_table_id = 1
        .deals_view = "deals_v"
        
        .employee_roles_table = "stella_common.employee_roles_t"
        .employee_role_links_table = "stella_common.employee_role_links_t"
        .entities_table = "stella_common.entities_t"
        
        .financial_advisors_table = "financial_advisors_t"
        .financial_advisors_view = "financial_advisors_v"
        .global_deal_list_view = "stella_common.global_deal_list_v"
        .global_policy_list_view = "stella_common.global_policy_list_v"
        .insurers_table = "stella_common.insurers_t"
        .insurers_view = "stella_common.insurers_v"
        .law_firms_table = "law_firms_t"
        .law_firms_view = "law_firms_v"
        .lawyers_table = "lawyers_t"
        .lawyers_view = "lawyers_v"
        .log_activity_table = "log_activity_t"
        .log_errors_eur_t = "stella_eur.log_errors_t"
        .log_errors_us_t = "stella_us.log_errors_t"
        .log_errors_v = "stella_common.log_errors_v"
        .jurisdictions_table = "stella_common.jurisdictions_t"
        .jurisdictions_view = "stella_common.jurisdictions_v"
        .menu_list_table = "stella_common.menu_list_t"
        
        .pricing_guidelines_table = "stella_common.pricing_t"
        .pricing_guidelines_view = "stella_common.pricing_v"
        .sectors_table = "stella_common.sectors_t"
        .service_messages_table = "stella_common.service_messages_t"
        .service_messages_view = "stella_common.service_messages_v"
        .templates_table = "stella_common.templates_t"
        .uws_table = "stella_common.underwriters_t"
        .uws_view = "stella_common.underwriters_v"
        .uw_role_fte_by_year_view = "stella_common.uw_role_fte_by_year_v"
        .uw_roles_view = "stella_common.uw_all_roles_v"
        
        .menu_lists.init
    End With
End Sub
Private Sub Class_Initialize()
    Me.init
End Sub