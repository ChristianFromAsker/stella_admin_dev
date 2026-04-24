Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public brands As String
Public broker_firms As String
Public broker_persons As String
Public budget_homes As String
Public country_list As String
Public deal_roles As String
Public deal_statuses As String
Public employee_roles As String
Public layer_types As String
Public law_firms As String
Public law_firms_uw_counsel As String
Public risk_types As String
Public row_source_brands_parties As String
Public row_source_budget_homes As String
Public row_source_deal_roles As String
Public row_source_deal_statuses As String
Public row_source_jurisdictions As String
Public row_source_risk_types As String
Public row_source_risk_types_reverse As String
Public row_source_sectors As String
Public row_source_uws As String
Public row_source_uws_current As String
Public stages As String
Public super_sectors As String
Public us_states As String
Public uws As String
Public yes_no As String
Public yes_no_maybe As String
Public uws_that_have_quote_righst As String
Public uws_that_have_binding_rights As String

Private Const int_yes As Integer = menu_list.yes
Private Const int_no As Integer = menu_list.no

Public Sub init()
    Dim proc_name As String
    proc_name = "cls_sources__menu_lists.init"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    
    Dim menu_item As String
    menu_item = "menu_item"
    If Load.system_info.app_continent = Load.system_info.continents.americas Then menu_item = "menu_item_us menu_item"
    
    With Me
        .brands = "SELECT brand_id id, brand_name menu_item" _
        & " FROM " & Load.sources.brands_table & " WHERE is_deleted = 0 AND is_party = 1 ORDER BY brand_name"
        
        .broker_firms = "SELECT broker_id id, broker_name menu_item FROM " & Load.sources.broker_firms_view & " ORDER BY short_name"

        .broker_persons = "SELECT broker_person_id id, personal_name menu_item FROM " & Load.sources.broker_persons_view & " ORDER BY personal_name"

        .budget_homes = "SELECT entity_id id, entity_business_name menu_item FROM " & Load.sources.entities_table _
        & " WHERE entity_type = 475 AND is_deleted = 0 ORDER BY entity_business_name"

        .country_list = "SELECT jurisdiction_id id, jurisdiction menu_item FROM " & Load.sources.jurisdictions_table _
        & " WHERE jurisdiction_type = 'country' AND is_deleted = 0 ORDER BY jurisdiction"
        
        .deal_roles = "SELECT deal_role_id id, deal_role_name menu_item" _
        & " FROM " & Load.sources.deal_roles_table & " WHERE is_deleted = 0 ORDER BY deal_role_name"
        
        .deal_statuses = "SELECT menu_id id, " & menu_item & " FROM " & Load.sources.menu_list_table & " WHERE item_type = 'Status' ORDER BY Setting2"
        
        .employee_roles = "SELECT employee_role_id id , employee_role_name menu_item FROM " & Load.sources.employee_roles_table _
        & " WHERE NOT parent_employee_role_id IS NULL AND is_deleted = 0"
        
        .law_firms = "SELECT law_firm_id id, firm_name menu_item FROM " & Load.sources.law_firms_view & " ORDER BY firm_name"
        
        .law_firms_uw_counsel = "SELECT law_firm_id id, firm_name menu_item FROM " _
        & Load.sources.law_firms_view _
        & " WHERE is_counsel = 93" _
        & " ORDER BY firm_name"
        
        .layer_types = "SELECT menu_id id, " & menu_item & " FROM " & Load.sources.menu_list_table & " WHERE item_type = 'layer_type'"

        .risk_types = "SELECT menu_id id, " & menu_item & " FROM " & Load.sources.menu_list_table & " WHERE item_type = 'RiskType'"
        
        .stages = "SELECT menu_id id, " & menu_item & " FROM " & Load.sources.menu_list_table & " WHERE item_type = 'stage'"
        
        .super_sectors = "SELECT sector_id id, sector_name menu_item FROM " & Load.sources.sectors_table _
        & " WHERE is_deleted = 0 AND sector_type = 494 AND NOT sector_id = 111" _
        & " ORDER BY sector_name ASC"
        
        .us_states = "SELECT jurisdiction_id id, jurisdiction menu_item FROM " & Load.sources.jurisdictions_view & " WHERE jurisdiction_type = 'us_state' ORDER BY jurisdiction"

        .uws = "SELECT id, uw_initials menu_item, uw_name menu_info FROM " & sources.underwriters_view _
        & " WHERE is_employed_id = 93 AND user_type = 150 ORDER BY uw_initials"

        .uws_that_have_quote_righst = "SELECT uw_id id, uw_initials menu_item, uw_name menu_info FROM " & Load.sources.approval_rights_view _
        & " WHERE has_nbi_approval_rights_id = " & int_yes & " ORDER BY uw_name"

        .uws_that_have_binding_rights = "SELECT uw_id id, uw_initials menu_item, uw_name menu_info FROM " & Load.sources.approval_rights_view _
        & " WHERE has_inception_approval_rights_id = " & int_yes & " ORDER BY uw_name"

        .yes_no = "SELECT menu_id id, " & menu_item & " FROM " & Load.sources.menu_list_table _
        & " WHERE item_type = 'YesNo' AND is_deleted = 0 AND is_active = 93 AND NOT menu_id = 182 AND NOT menu_id = 460"

        .yes_no_maybe = "SELECT menu_id id, " & menu_item & " FROM " & Load.sources.menu_list_table & " WHERE item_type = 'YesNo' AND is_deleted = 0 AND is_active = 93"

    End With
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
        '. Open
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
        '. Open
        Do Until .EOF
            If output <> "" Then output = output & ";"
            output = output & !id & ";'" & !menu_item & "'"
            .MoveNext
        Loop
        .Close
    End With
    Me.row_source_sectors = output
    
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