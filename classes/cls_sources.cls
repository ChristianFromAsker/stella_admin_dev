Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public add_many_deals_f As String
Public approval_rights_view As String
Public awac_excels_table As String
Public awac_excels_view As String

Public binder_limits_table As String
Public binder_limits_view As String
Public binder_list_table As String
Public binder_list_view As String
Public binder_jurisdictions_table As String
Public binder_jurisdictions_view As String
Public binder_questions_table As String
Public binder_questions_view As String
Public binder_risk_types_table As String
Public binder_risk_types_view As String

Public brands_table As String
Public brands_view As String
Public broker_firms_table As String
Public broker_firms_view As String
Public broker_persons_table As String
Public broker_persons_view As String
Public budget_homes_view As String

Public claims_table As String
Public cm_binder_questions_with_bad_sectors_view As String
Public cm_country_procedures_templates_table As String
Public cm_country_procedures_templates_view As String
Public cm_country_procedures_deals_table As String
Public cm_country_procedures_deals_view As String
Public cm_country_procedures_statuses_table As String
Public cm_key_data_table As String
Public cm_policies_table As String
Public currencies_table As String

Public data_fields_table As String
Public deal_parties_table As String
Public deal_parties_view As String
Public deal_questions_table As String
Public deal_questions_view As String
Public deal_roles_table As String
Public deal_statuses_view As String
Public deal_tags_table As String
Public deals_for_auto_archiving_view As String
Public deals_table As String
Public deals_view As String
Public duplicate_check_view As String
Public dupliacte_deals_view As String

Public employee_role_links_table As String
Public employee_roles_table As String
Public entities_table As String
Public extra_deal_limit_table As String
Public extra_deal_limit_view As String

Public financial_advisors_table As String
Public financial_advisors_view As String
Public folder_moving_table As String

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
Public live_deals_view As String
Public log_activity_table As String
Public log_data_table As String
Public log_errors_eur_t As String
Public log_errors_us_t As String
Public log_errors_v As String

Public master_deal_list_view As String
Public menu_list_table As String
Public my_deals_view As String

Public navins_homes_table As String

Public parties_table As String
Public policies_table As String
Public policies_table_id As String
Public policies_view As String
Public policy_binders_table As String
Public policy_binders_table_id As Integer
Public policy_binders_view As String
Public policy_extract_for_uw_view As String
Public policy_layer_texts_table As String
Public pricing_guidelines_table As String
Public pricing_guidelines_view As String

Public referrals_table As String
Public referrals_view As String
Public risk_details_f_view As String
Public risk_types_table As String
Public rp_entity_info_view As String

Public sectors_table As String
Public sectors_view As String
Public service_messages_table As String
Public service_messages_view As String
Public submission_control_view As String

Public template_deal_tags_table As String
Public template_qs_table As String
Public template_qs_view As String

Public underwriters_table As String
Public underwriters_view As String
Public uw_authority_table As String
Public uw_authority_view As String
Public uw_positions_view As String
Public uw_financial_persons_view As String
Public uw_role_fte_by_year_view As String
Public uw_roles_view As String

Public yes_no_view As String

Public weekly_view_f As String


Public menu_lists As New cls_sources__menu_lists
Public is_init As Boolean

Public Sub init()
    Const proc_name As String = "cls_sources.init"
    utilities.call_stack_add_item proc_name
    
    With Me
        
        .add_many_deals_f = "recent_deals_v"
        .approval_rights_view = "approval_rights_v"
        .awac_excels_table = "stella_common.awac_excels_t"
        .awac_excels_view = "stella_common.awac_excels_v"
        
        .binder_jurisdictions_table = "stella_common.cm_binder_jurisdictions_t"
        .binder_jurisdictions_view = "stella_common.cm_binder_jurisdictions_v"
        .binder_limits_table = "stella_common.cm_binder_limits_t"
        .binder_list_table = "stella_common.binders_t"
        .binder_list_view = "stella_common.binders_v"
        .binder_questions_table = "stella_common.cm_binder_questions_t"
        .binder_questions_view = "stella_common.cm_binder_questions_v"
        .binder_risk_types_table = "stella_common.cm_binder_risk_types_t"
        .binder_risk_types_view = "stella_common.cm_binder_risk_types_v"
        .brands_table = "stella_common.brands_t"
        .brands_view = "stella_common.brands_v"
        .broker_firms_table = "broker_firms_t"
        .broker_firms_view = "broker_firms_v"
        .broker_persons_table = "broker_persons_t"
        .broker_persons_view = "broker_persons_v"
        .budget_homes_view = "stella_common.budget_homes_v"
        
        .cm_binder_questions_with_bad_sectors_view = "cm_binder_questions_with_bad_sectors_v"
        .cm_country_procedures_deals_table = "cm_country_procedures_deals_t"
        .cm_country_procedures_deals_view = "cm_country_procedures_deals_v"
        .cm_country_procedures_templates_table = "stella_common.cm_country_procedures_templates_t"
        .cm_country_procedures_templates_view = "stella_common.cm_country_procedures_templates_v"
        .cm_country_procedures_statuses_table = "stella_common.cm_country_procedures_statuses_t"
        .cm_key_data_table = "cm_deals_t"
        .cm_policies_table = "cm_policies_t"
            
        .claims_table = "stella_eur.claims_t"
        .currencies_table = "stella_common.currencies_t"
        
        .data_fields_table = "data_fields_t"
        .deal_parties_table = "deal_parties_t"
        .deal_parties_view = "deal_parties_v"
        .deal_questions_table = "cm_deal_questions_t"
        .deal_questions_view = "cm_deal_questions_v"
        .deal_roles_table = "stella_common.deal_roles_t"
        .deal_statuses_view = "stella_common.deal_statuses_v"
        .deal_tags_table = "deal_tags_t"
        .deals_for_auto_archiving_view = "deals_for_auto_archiving_v"
        .deals_table = "deals_t"
        .deals_view = "deals_v"
        .dupliacte_deals_view = "duplicate_deals_v"
        .duplicate_check_view = "deals_for_duplicate_check_v"
        
        .employee_role_links_table = "stella_common.employee_role_links_t"
        .employee_roles_table = "stella_common.employee_roles_t"
        .entities_table = "stella_common.entities_t"
        .extra_deal_limit_table = "stella_us.cm_extra_deal_limit_t"
        .extra_deal_limit_view = "stella_us.cm_extra_deal_limit_v"
        
        .financial_advisors_table = "financial_advisors_t"
        .financial_advisors_view = "financial_advisors_v"
        .folder_moving_table = "folder_moving_t"
    
        .global_deal_list_view = "stella_common.global_deal_list_v"
        .global_policy_list_view = "stella_common.global_policy_list_v"
        
        .insurers_table = "stella_common.insurers_t"
        .insurers_view = "stella_common.insurers_v"
        
        .jurisdictions_table = "stella_common.jurisdictions_t"
        .jurisdictions_view = "stella_common.jurisdictions_v"
        
        .law_firms_table = "law_firms_t"
        .law_firms_view = "law_firms_v"
        .lawyers_table = "lawyers_t"
        .lawyers_view = "lawyers_v"
        .live_deals_view = "live_deals_v"
        .log_activity_table = "log_activity_t"
        .log_data_table = "log_data_t"
        .log_errors_eur_t = "stella_eur.log_errors_t"
        .log_errors_us_t = "stella_us.log_errors_t"
        .log_errors_v = "stella_common.log_errors_v"
        
        .master_deal_list_view = "master_deal_list_v"
        .menu_list_table = "stella_common.menu_list_t"
        .my_deals_view = "my_deals_v"
        
        .navins_homes_table = "stella_common.navins_homes_t"
        
        .parties_table = "parties_t"
        .policy_binders_table = "security_t"
        .policy_binders_view = "security_v"
        .policy_layer_texts_table = "stella_common.policy_layer_texts_t"
        .policies_table = "layers_t"
        .policies_view = "policies_v"
        .pricing_guidelines_table = "stella_common.pricing_t"
        .pricing_guidelines_view = "stella_common.pricing_v"
        
        .referrals_table = "cm_referrals_t"
        .referrals_view = "cm_referrals_v"
        .risk_details_f_view = "risk_details_f_v"
        .risk_types_table = "stella_common.risk_types_t"
        .rp_entity_info_view = "stella_common.rp_entity_info_v"
        
        .sectors_table = "stella_common.sectors_t"
        .sectors_view = "stella_common.sectors_v"
        .service_messages_table = "stella_common.service_messages_t"
        .service_messages_view = "stella_common.service_messages_v"
        .submission_control_view = "sub_control_v"
        
        .template_deal_tags_table = "stella_common.template_deal_tags_t"
        .template_qs_table = "stella_common.cm_template_questions_t"
        .template_qs_view = "stella_common.cm_template_questions_v"
        
        .underwriters_table = "stella_common.underwriters_t"
        .underwriters_view = "stella_common.underwriters_v"
        .uw_financial_persons_view = "uw_financial_persons_v"
        .uw_positions_view = "cm_uw_positions_v"
        .uw_role_fte_by_year_view = "stella_common.uw_role_fte_by_year_v"
        .uw_roles_view = "stella_common.uw_all_roles_v"
        
        .weekly_view_f = "weekly_view_v"
        
        .menu_lists.init
    End With
    
    is_init = True
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
End Sub