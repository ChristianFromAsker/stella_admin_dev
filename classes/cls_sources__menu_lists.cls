Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public brands As String
Public broker_firms As String
Public broker_persons As String
Public budget_continents As String
Public budget_homes As String
Public budget_regions As String
Public deal_roles As String
Public deal_statuses As String
Public employee_products As String
Public employee_roles As String
Public layer_types As String
Public law_firms As String
Public risk_types As String
Public risk_types_major As String
Public risk_types_incl_major As String
Public row_source_brands_parties As String
Public row_source_budget_homes As String
Public row_source_deal_roles As String
Public row_source_deal_statuses As String
Public row_source_employee_products As String
Public row_source_employee_roles As String
Public row_source_jurisdictions As String
Public row_source_risk_types As String
Public row_source_risk_types_reverse As String
Public row_source_sectors As String
Public row_source_template_deal_tags As String
Public row_source_uws As String
Public row_source_uws_current As String
Public stages As String
Public super_sectors As String
Public template_deal_tags As String
Public us_states As String
Public uws As String
Public yes_no As String
Public uws_that_have_quote_righst As String
Public uws_that_have_binding_rights As String

Public Sub init()
    Const proc_name As String = "cls_sources__menu_lists"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0

    Dim menu_item As String
    Dim risk_type_item As String
    
    menu_item = "menu_item"
    If Load.system_info.app_continent = Load.system_info.continents.americas Then menu_item = "menu_item_us menu_item"
    
    With Me
        .brands = "SELECT brand_id id, brand_name menu_item" _
        & " FROM " & Load.sources.brands_table & " WHERE is_deleted = 0 AND is_party = 1 ORDER BY brand_name"
        
        .broker_firms = "SELECT broker_firm_id id, business_name menu_item" _
        & " FROM " & Load.sources.broker_firms_view _
        & " ORDER BY business_name"
        
        .broker_persons = "SELECT broker_person_id id, personal_name menu_item FROM " & Load.sources.broker_persons_view & " ORDER BY personal_name"
        
        .budget_continents = "SELECT jurisdiction_id id, jurisdiction menu_item FROM " _
        & Load.sources.jurisdictions_table _
        & " WHERE is_deleted = 0 AND jurisdiction_type = 'continent'"
        
        .budget_homes = "SELECT budget_home_id id, budget_home menu_item FROM " & Load.sources.budget_homes_view _
        & " ORDER BY menu_item ASC"
        
        .budget_regions = "SELECT DISTINCT(budget_region) menu_item, budget_region_id id FROM " & Load.sources.budget_homes_view
        
        .deal_roles = "SELECT deal_role_id id, deal_role_name menu_item" _
        & " FROM " & Load.sources.deal_roles_table & " WHERE is_deleted = 0 ORDER BY deal_role_name"
        
        .deal_statuses = "SELECT menu_id id, " & menu_item & " FROM " & Load.sources.deal_statuses_view & " WHERE status_type = 'deal_status' ORDER BY order_by"
        
        .employee_products = "SELECT employee_product_id id, employee_product_name menu_item" _
        & " FROM " & Load.sources.employee_products_table _
        & " WHERE is_deleted = 0 AND parent_employee_product_id__employee_products_t > 0" _
        & " ORDER BY employee_product_name ASC"
        
        .employee_roles = "SELECT employee_role_id id, employee_role_name menu_item" _
        & " FROM " & Load.sources.employee_roles_table _
        & " WHERE is_deleted = 0 AND parent_employee_role_id > 0" _
        & " ORDER BY employee_role_name ASC"
        
        .law_firms = "SELECT law_firm_id id, firm_name menu_item FROM " & Load.sources.law_firms_view & " ORDER BY firm_name"
        
        .layer_types = "SELECT menu_id id, " & menu_item _
        & " FROM " & Load.sources.menu_list_table _
        & " WHERE item_type = 'layer_type' AND is_deleted = 0"
        
        risk_type_item = "risk_type_name menu_item"
        If Load.system_info.app_continent = Load.system_info.continents.americas Then
            risk_type_item = "risk_type_name_us menu_item"
        End If
        
        .risk_types_incl_major = "SELECT risk_type_id id," & risk_type_item & " FROM " & Load.sources.risk_types_table _
        & " WHERE is_deleted = 0 AND show_in_filters_id = 93" _
        & " ORDER BY risk_type_name"
        
        .risk_types = "SELECT risk_type_id id, " & risk_type_item & " FROM " & Load.sources.risk_types_table _
        & " WHERE is_deleted = 0 AND NOT risk_type_major_id__risk_types_t IS NULL" _
        & " ORDER BY risk_type_name"
        
        .risk_types_major = "SELECT risk_type_id id, " & risk_type_item & " FROM " & Load.sources.risk_types_table _
        & " WHERE is_deleted = 0 AND risk_type_major_id__risk_types_t IS NULL"
        
        .stages = "SELECT menu_id id, " & menu_item & " FROM " & Load.sources.menu_list_table & " WHERE item_type = 'stage'"
        
        .super_sectors = "SELECT sector_id id, sector_name menu_item FROM " & Load.sources.sectors_table _
        & " WHERE is_deleted = 0 AND sector_type = 494 AND NOT sector_id = 111" _
        & " ORDER BY sector_name ASC"
        
        .template_deal_tags = "SELECT tag_id id, tag_name menu_item FROM " & Load.sources.template_deal_tags_table & " WHERE is_deleted = 0 ORDER BY tag_name"
        
        .us_states = "SELECT jurisdiction_id id, jurisdiction menu_item FROM " & Load.sources.jurisdictions_view _
        & " WHERE jurisdiction_type = 'us_state' ORDER BY jurisdiction"
        
        .uws = "SELECT uw_id id, CONCAT(uw_initials, ' (', uw_name, ')') menu_item FROM " & sources.underwriters_view _
        & " WHERE is_employed_id = 93 AND user_type = 150 ORDER BY uw_initials"
        
        .uws_that_have_quote_righst = "SELECT uw_id id, CONCAT(uw_initials, ' (', uw_name, ')') menu_item" _
        & " FROM " & Load.sources.approval_rights_view _
        & " WHERE has_nbi_approval_rights_id = " & Load.yes & " ORDER BY uw_name"
        
        .uws_that_have_binding_rights = "SELECT uw_id id, CONCAT(uw_initials, ' (', uw_name, ')') menu_item" _
        & " FROM " & Load.sources.approval_rights_view _
        & " WHERE has_inception_approval_rights_id = " & Load.yes & " ORDER BY uw_name"
        
        .yes_no = "SELECT menu_id id, " & menu_item & " FROM " & Load.sources.menu_list_table & " WHERE item_type = 'YesNo' AND is_deleted = 0 AND is_active = 93"
    End With
outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub

Public Sub init_row_sources()
    Const proc_name As String = "cls_sources__menu_lists.init_row_soruce_country_list"
    utilities.call_stack_add_item proc_name
    
    Dim i As Long
    Dim country_id As Long
    Dim country_name As String
    Dim output As String
    Dim output_2 As String
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    Dim uw_id As Long
    Dim uw_initials As String
    Dim uw_is_employed_id As Long
    Dim uw_name As String
    Dim uw_user_type_id As Long
    
    'brands
    output = ""
    str_sql = Me.brands
    Set rs = utilities.create_adodb_rs(Load.conn, str_sql)
    With rs
        Do Until .EOF
            If output <> "" Then output = output & ";"
            output = output & !id & ";'" & !menu_item & "'"
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Me.row_source_brands_parties = output
    
    'budget homes
    output = ""
    str_sql = Me.budget_homes
    Set rs = utilities.create_adodb_rs(Load.conn, str_sql)
    With rs
        '. Open
        Do Until .EOF
            If output <> "" Then output = output & ";"
            output = output & !id & ";'" & !menu_item & "'"
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Me.row_source_budget_homes = output
    
    'country list
    output = ""
    For i = 1 To Load.country_list(0, 0)
        country_id = Load.country_list(i, 0)
        country_name = Load.country_list(i, 1)
        If output <> "" Then output = output & ";"
        output = output & country_id & ";'" & country_name & "'"
    Next i
    Me.row_source_jurisdictions = output
    
    'deal roles
    output = ""
    str_sql = Me.deal_roles
    Set rs = utilities.create_adodb_rs(Load.conn, str_sql)
    With rs
        '. Open
        Do Until .EOF
            If output <> "" Then output = output & ";"
            output = output & !id & ";'" & !menu_item & "'"
            .MoveNext
        Loop
        .Close
    End With
    Me.row_source_deal_roles = output
    
    'deal status list
    output = ""
    str_sql = Me.deal_statuses
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        '. Open
        Do Until .EOF
            If output <> "" Then output = output & ";"
            output = output & !id & ";'" & !menu_item & "'"
            .MoveNext
        Loop
        .Close
    End With
    Me.row_source_deal_statuses = output
    
    'employee products
    output = ""
    str_sql = Me.employee_products
    Set rs = utilities.create_adodb_rs(Load.conn, str_sql)
    With rs
        Do Until .EOF
            If output <> "" Then output = output & ";"
            output = output & !id & ";'" & !menu_item & "'"
            .MoveNext
        Loop
        Me.row_source_employee_products = output
        .Close
    End With
    
    'employee roles
    output = ""
    str_sql = Me.employee_roles
    Set rs = utilities.create_adodb_rs(Load.conn, str_sql)
    With rs
        Do Until .EOF
            If output <> "" Then output = output & ";"
            output = output & !id & ";'" & !menu_item & "'"
            .MoveNext
        Loop
        Me.row_source_employee_roles = output
        .Close
    End With
    
    'risk types
    output = ""
    str_sql = Me.risk_types
    Set rs = utilities.create_adodb_rs(Load.conn, str_sql)
    With rs
        Do Until .EOF
            If output <> "" Then output = output & ";"
            output = output & !id & ";'" & !menu_item & "'"
            .MoveNext
        Loop
        Me.row_source_risk_types = output
        output = ""
        .MovePrevious
        Do Until .BOF
            If output <> "" Then output = output & ";"
            output = output & !id & ";'" & !menu_item & "'"
            .MovePrevious
        Loop
        Me.row_source_risk_types_reverse = output
        .Close
    End With
    
    'sectors
    output = ""
    str_sql = Me.super_sectors
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        Do Until .EOF
            If output <> "" Then output = output & ";"
            output = output & !id & ";'" & !menu_item & "'"
            .MoveNext
        Loop
        .Close
    End With
    Me.row_source_sectors = output
    
    'template deal tags
    output = ""
    str_sql = Me.template_deal_tags
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        Do Until .EOF
            If output <> "" Then output = output & ";"
            output = output & !id & ";'" & !menu_item & "'"
            .MoveNext
        Loop
        .Close
    End With
    Me.row_source_template_deal_tags = output
    
    'uw list
    output = ""
    output_2 = ""
    uw_user_type_id = 150
    For i = 1 To Load.underwriters(0, 0)
        If Load.underwriters(i, 16) = uw_user_type_id Then
            uw_id = Load.underwriters(i, 1)
            uw_initials = Load.underwriters(i, 2)
            uw_name = Load.underwriters(i, 3)
            If output <> "" Then output = output & ";"
            output = output & uw_id & ";'" & uw_initials & "  (" & uw_name & ")'"
            If Load.underwriters(i, 15) = menu_list.yes Then
                If output_2 <> "" Then output_2 = output_2 & ";"
                output_2 = output_2 & uw_id & ";'" & uw_initials & "  (" & uw_name & ")'"
            End If
        End If
    Next i
    Me.row_source_uws = output
    Me.row_source_uws_current = output_2
    
outro:
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub