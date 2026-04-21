Option Compare Database
Option Explicit

'global variables
    'collections
    Public col_rp_entities As New Collection

    'form objects
    Public awac_excels As New cls_form_awac_excels
    Public awac_excels_new As New cls_form_awac_excels
    Public system_dashboard As New cls_form_system_dashboard
    
    'types
    Private Type typ_awac_excels_binders
        awac_can As String
        awac_usa As String
    End Type
    Public awac_excels_binders As typ_awac_excels_binders
    
    Private Type typ_awac_excels_folders
        awac_can__archive_output As String
        awac_can__archive_source As String
        awac_can__archive_test_output As String
        awac_can__archive_test_source As String
        awac_usa__archive_output As String
        awac_usa__archive_source As String
        awac_usa__archive_test_output As String
        awac_usa__archive_test_source As String
    End Type
    Public awac_excels_folders As typ_awac_excels_folders
    
    Private Type typ_control_list_properties
        field_name As Integer
        is_mandatory As Integer
        set_bg_to_white As Integer
        field_in_recordset As Integer
        empty_value As Integer
        control_object As Integer
    End Type
    Public control_list_propertes As typ_control_list_properties
    
    Private Type typ_field_properties
        field_name As Integer
        is_mandatory As Integer
        shall_reset As Integer
        item_count As Integer
    End Type
    Public field_properties As typ_field_properties

    Public Type typ_field_type
        cmd_button As String
        label As String
        text_field As String
        header As String
    End Type
    Public field_type As global_vars.typ_field_type
    
Public Sub init()
    Const proc_name As String = "global_vars.init"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Load.colors.init
    
    global_vars.init_awac_excels_vars
    global_vars.init_cols
    global_vars.init_variables_types
    
    'forms
    global_vars.awac_excels.init
    global_vars.system_dashboard.init
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub init_conn_dependant()
    Const proc_name As String = "global_vars.init_conn_dependant"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    If global_policy_list.is_init = False Then
        global_policy_list.init
    End If
    
    Load.init_array_underwriters
    Load.init_country_list_array
    Load.sources.menu_lists.init_row_sources
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub init_cols()
    Const proc_name As String = "global_vars.init_cols"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    global_vars.init_cols_rp_entities
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Sub init_cols_rp_entities()
    Const proc_name As String = "global_vars.init_cols_rp_entities"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim ctrl As cls_field
    Dim dict_rp_entity As Scripting.Dictionary
    Dim rs As ADODB.Recordset
    Dim str_sql As String
        
    str_sql = "SELECT entity_id, entity_business_name" _
    & " FROM " & Load.sources.entities_table _
    & " WHERE entity_type = 475" _
    & " ORDER BY entity_business_name ASC"
    
    Set rs = utilities.create_adodb_rs(Load.conn, str_sql)
    With rs
        Do Until .EOF
            Set ctrl = New cls_field
            ctrl.id = !entity_id.Value
            ctrl.field_caption = !entity_business_name.Value
            global_vars.col_rp_entities.Add ctrl
            .MoveNext
        Loop
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
Public Sub init_variables_types()
    With global_vars.control_list_propertes
        .field_name = 0
        .is_mandatory = 1
        .set_bg_to_white = 2
        .field_in_recordset = 3
        .empty_value = 4
        .control_object = 5
    End With
    With global_vars.field_properties
        .field_name = 0
        .is_mandatory = 1
        .shall_reset = 2
        .item_count = 3
    End With
    With global_vars.field_type
        .cmd_button = "command_button"
        .header = "header"
        .label = "label"
        .text_field = "text_field"
    End With
End Sub

Public Sub init_awac_excels_vars()
    With global_vars.awac_excels_folders
        .awac_can__archive_output = "awac_can__archive\reports_output\"
        .awac_can__archive_source = "awac_can__archive\reports_source\"
        .awac_can__archive_test_output = "awac_can__archive__test\reports_output\"
        .awac_can__archive_test_source = "awac_can__archive__test\reports_source\"
        .awac_usa__archive_output = "awac_usa__archive\reports_output\"
        .awac_usa__archive_source = "awac_usa__archive\reports_source\"
        .awac_usa__archive_test_output = "awac_usa__archive__test\reports_output\"
        .awac_usa__archive_test_source = "awac_usa__archive__test\reports_source\"
    End With
    
    'Tuesday 14 Jan 2025, CK: These are the carrier account codes (RP0001), found in colimn D in the V5 written reports
    With global_vars.awac_excels_binders
        .awac_can = "2738"
        .awac_usa = "2638"
    End With
End Sub