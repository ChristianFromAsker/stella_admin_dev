Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'system elements
Public form_name As String
Public make_all_fields_visible_for_debugging As Boolean
Public error_excel_conversion As Boolean

'design variables
Public header_lbl_email As New cls_field
Public header_txt_email As New cls_field
Public header_lbl_to_field As New cls_field
Public header_txt_to_field As New cls_field
Public header_lbl_cc_field As New cls_field
Public header_txt_cc_field As New cls_field
Public header_lbl_payment_report As New cls_field
Public header_txt_payment_report As New cls_field
Public header_lbl_v5_report As New cls_field
Public header_txt_v5_report As New cls_field
Public header_lbl_awac_id As New cls_field
Public header_txt_awac_id As New cls_field

Public header_cmd_add_change As New cls_field
Public header_cmd_emails As New cls_field
Public header_cmd_goto_folder As New cls_field
Public header_cmd_prepare_excels As New cls_field
Public header_cmd_prepare_excels_test As New cls_field
Public header_cmd_show_input As New cls_field

Public header_instructions As New cls_field

Public col_input_fields As New Collection
Public col_all_fields As New Collection
Public col_cmds As New Collection

'excel generation variables
Public col_file_list_for_email_attaching As New Collection
Public current_output_folder As String
Public current_source_folder As String
Public current_source_report_type As String
Public found_payment_report As Boolean
Public found_v5_report As Boolean
Public reports As New cls_form_awac_excels_reports

Public Sub init_collections()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_forms_awac_excels.init_collections"
    Dim form_field As cls_field
    
    Set Me.col_input_fields = Nothing
    With Me.col_input_fields
        .Add Me.header_lbl_awac_id
        .Add Me.header_txt_awac_id
        .Add Me.header_lbl_cc_field
        .Add Me.header_txt_cc_field
        .Add Me.header_lbl_email
        .Add Me.header_txt_email
        .Add Me.header_lbl_payment_report
        .Add Me.header_txt_payment_report
        .Add Me.header_lbl_to_field
        .Add Me.header_txt_to_field
        .Add Me.header_lbl_v5_report
        .Add Me.header_txt_v5_report
        .Add Me.header_cmd_add_change
    End With
    
    Set Me.col_cmds = Nothing
    With Me.col_cmds
        .Add Me.header_cmd_add_change
        .Add Me.header_cmd_emails
        .Add Me.header_cmd_goto_folder
        .Add Me.header_cmd_prepare_excels
        .Add Me.header_cmd_prepare_excels_test
        .Add Me.header_cmd_show_input
    End With
    
    Set Me.col_all_fields = Nothing
    With Me.col_all_fields
        For Each form_field In Me.col_input_fields
            .Add form_field
        Next form_field
        For Each form_field In Me.col_cmds
            .Add form_field
        Next form_field
        .Add Me.header_instructions
    End With
End Sub

Public Sub init()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_forms_awac_excels.init"
    Dim form_field As cls_field
    Me.form_name = "awac_excels_f"
    Me.init_collections
    
    For Each form_field In Me.col_input_fields
        With form_field
            .field_visible = False
            .field_left = 100
            .field_height = 340
            .field_width = 1800
        End With
    Next form_field
        
    With Me.header_lbl_email
        .field_name = "header_lbl_email"
        .field_top = 100
        Set form_field = Me.header_lbl_email
    End With
    With Me.header_txt_email
        .field_name = "header_email"
        .field_left = Me.header_lbl_email.field_right
        .field_top = Me.header_lbl_email.field_top
    End With
    
    With Me.header_lbl_to_field
        .field_name = "header_lbl_to_field"
        .field_top = form_field.field_bottom
        Set form_field = Me.header_lbl_to_field
    End With
    With Me.header_txt_to_field
        .field_name = "header_to_field"
        .field_left = Me.header_lbl_to_field.field_right
        .field_top = Me.header_lbl_to_field.field_top
    End With
    
    With Me.header_lbl_cc_field
        .field_name = "header_lbl_cc_field"
        .field_top = form_field.field_bottom
        Set form_field = Me.header_lbl_cc_field
    End With
    With Me.header_txt_cc_field
        .field_name = "header_cc_field"
        .field_left = Me.header_lbl_cc_field.field_right
        .field_top = Me.header_lbl_cc_field.field_top
    End With
    
    With Me.header_lbl_payment_report
        .field_name = "header_lbl_payment_report"
        .field_top = form_field.field_bottom
        Set form_field = Me.header_lbl_payment_report
    End With
    With Me.header_txt_payment_report
        .field_name = "header_payment_report"
        .field_left = Me.header_lbl_payment_report.field_right
        .field_top = Me.header_lbl_payment_report.field_top
    End With
    
    With Me.header_lbl_v5_report
        .field_name = "header_lbl_v5_report"
        .field_top = form_field.field_bottom
        Set form_field = Me.header_lbl_v5_report
    End With
    With Me.header_txt_v5_report
        .field_name = "header_v5_report"
        .field_left = Me.header_lbl_v5_report.field_right
        .field_top = Me.header_lbl_v5_report.field_top
    End With
    
    With Me.header_lbl_awac_id
        .field_name = "header_lbl_deal_id"
        .field_top = form_field.field_bottom
        Set form_field = Me.header_lbl_awac_id
    End With
    With Me.header_txt_awac_id
        .field_name = "header_awac_id"
        .field_left = Me.header_lbl_awac_id.field_right
        .field_top = Me.header_lbl_awac_id.field_top
    End With
    
    With Me.header_instructions
        .field_height = utilities.twips_converter(5, "cm")
        .field_name = "header_instructions"
        .field_visible = True
        .field_width = 4500
        .field_top = 100
        .field_left = 100
    End With
    
    'cmds
    Dim cmd_height As Integer
    cmd_height = 600
    With Me.header_cmd_goto_folder
        .field_height = cmd_height
        .field_name = "cmd_goto_folder"
        .field_visible = True
        .field_left = Me.header_instructions.field_right
        .field_top = Me.header_instructions.field_top
    End With
    With Me.header_cmd_add_change
        .field_height = cmd_height
        .field_name = "cmd_add_change"
        .field_visible = False
        .field_top = 100
        .field_left = 5000
    End With
    With Me.header_cmd_prepare_excels_test
        .field_height = cmd_height
        .field_name = "cmd_prepare_excels_test"
        .field_visible = True
        .field_top = 100
        .field_left = 8000
        Set form_field = Me.header_cmd_prepare_excels_test
    End With
    With Me.header_cmd_prepare_excels
        .field_height = cmd_height
        .field_name = "cmd_prepare_excels"
        .field_visible = True
        .field_top = form_field.field_bottom + 100
        .field_left = 8000
        Set form_field = Me.header_cmd_prepare_excels
    End With
    With Me.header_cmd_emails
        .field_height = cmd_height
        .field_name = "cmd_prepare_emails"
        .field_visible = True
        .field_top = form_field.field_bottom + 100
        .field_left = 8000
        Set form_field = Me.header_cmd_emails
    End With
    With Me.header_cmd_show_input
        .field_height = cmd_height
        .field_name = "cmd_show_input"
        .field_visible = True
        .field_top = form_field.field_bottom + 100
        .field_left = 8000
        Set form_field = Me.header_cmd_show_input
    End With
    Set form_field = Nothing
End Sub

Public Sub paint()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_forms_awac_excels.paint"
    Dim form_field As cls_field
    
    For Each form_field In Me.col_all_fields
        With Forms(Me.form_name).Controls(form_field.field_name)
            .Height = form_field.field_height
            .Visible = form_field.field_visible
            .Left = form_field.field_left
            .Top = form_field.field_top
        End With
    Next form_field
    Debug.Print "test"
End Sub

Public Sub show_hide_input()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_forms_awac_excels.show_hide_input"
    Dim form_field As cls_field, new_visible_status As Boolean, old_visible_status As Boolean
    Forms(Me.form_name)!email_address.SetFocus
    new_visible_status = False
    old_visible_status = True
    If Me.header_lbl_email.field_visible = False Then
        new_visible_status = True
        old_visible_status = False
    End If
    For Each form_field In Me.col_input_fields
        With form_field
            .field_visible = new_visible_status
        End With
    Next form_field
    Me.header_instructions.field_visible = old_visible_status
    Me.header_cmd_goto_folder.field_visible = old_visible_status
    Me.header_cmd_show_input.field_visible = old_visible_status
    Me.paint
    
End Sub

Public Function emails_convert_excels_binder_identification(ByVal report_data, ByVal is_test_run As Boolean) As String
    Dim proc_name As String
    proc_name = "cls_form_awac_excels.emails_convert_excels_binder_identification"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    If IsArray(report_data) = False Then
        emails_convert_excels_binder_identification = "-1"
        GoTo outro
    End If
    
    Dim binder_id As String
    Dim row_counter As Integer
    Dim column_counter As Integer
    Dim cell_content As String
    Dim header_row As Long
    
    binder_id = "-1"
    
    'find header row
    header_row = "-1"
    For row_counter = 1 To 20
        cell_content = report_data(row_counter, 1)
        If InStr(1, cell_content, "CR0001") > 0 Then
            header_row = row_counter
        End If
        If header_row <> "-1" Then row_counter = 20
    Next row_counter

    'find colum with 'carrier account code (RP0001)'
    binder_id = "-1"
    For column_counter = 1 To 30
        cell_content = report_data(header_row, column_counter)
        If InStr(1, cell_content, "RP0001") > 0 Then
            binder_id = CStr(CInt(report_data(header_row + 1, column_counter)))
        End If
        If binder_id <> "-1" Then column_counter = 30
    Next column_counter
    
    If binder_id = "-1" Then
        emails_convert_excels_binder_identification = "-1"
        GoTo outro
    End If
    
    With global_vars.awac_excels_folders
        If binder_id = global_vars.awac_excels_binders.awac_can Then
            emails_convert_excels_binder_identification = global_vars.awac_excels_binders.awac_can
            Me.current_source_folder = Load.system_info.system_paths.awac_conversion_path & .awac_can__archive_source
            Me.current_output_folder = Load.system_info.system_paths.awac_conversion_path & .awac_can__archive_output
            If is_test_run = True Then
                Me.current_source_folder = Load.system_info.system_paths.awac_conversion_path & .awac_can__archive_test_source
                Me.current_output_folder = Load.system_info.system_paths.awac_conversion_path & .awac_can__archive_test_output
            End If
        ElseIf binder_id = global_vars.awac_excels_binders.awac_usa Then
            Me.current_source_folder = Load.system_info.system_paths.awac_conversion_path & .awac_usa__archive_source
            Me.current_output_folder = Load.system_info.system_paths.awac_conversion_path & .awac_usa__archive_output
            If is_test_run = True Then
                Me.current_source_folder = Load.system_info.system_paths.awac_conversion_path & .awac_usa__archive_test_source
                Me.current_output_folder = Load.system_info.system_paths.awac_conversion_path & .awac_usa__archive_test_output
            End If
            emails_convert_excels_binder_identification = global_vars.awac_excels_binders.awac_usa
        Else
            Me.current_source_folder = "-1"
            Me.current_output_folder = "-1"
            emails_convert_excels_binder_identification = "-1"
        End If
    End With
outro:
    Exit Function

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Function
Public Sub emails_convert_excels(ByVal is_test_run As Boolean)
    Dim proc_name As String, str_milestone As String
    proc_name = "cls_form_awac_excels.emails_convert_excels"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    'there are two source excels, though with a old and new naming convetion. The code will identify them by their sheet names
    Dim v5_source_sheet As String
    Dim v5_source_sheet_old As String
    Dim payment_source_sheet As String
    Dim payment_source_sheet_old As String
    v5_source_sheet = "V5WrittenReport"
    v5_source_sheet_old = "V5 Written Report"
    payment_source_sheet = "V5PayableReport"
    payment_source_sheet_old = "V5 Payable Report"
    
    Me.found_v5_report = False
    Me.found_payment_report = False
    
    Dim app_excel As Excel.Application
    Dim awac_account_current_name_test As String
    Dim awac_payment_name_test As String
    Dim awac_v5_name_test As String
    Dim binder_currency As Variant
    Dim col_source_files As New Collection
    Dim column_counter As Long
    Dim currency_column As Integer
    Dim current_binder As String
    Dim do_next_file As Boolean
    Dim excel_data() As Variant
    Dim excel_file_for_email As Scripting.Dictionary
    Dim excel_row_counter As Long
    Dim file_mover As New FileSystemObject
    Dim file_name As String
    Dim file_name_component As String
    Dim file_name_component_2 As String
    Dim filename As String
    Dim folder_path As String
    Dim fso As Object
    Dim i As Integer
    Dim last_row As Long
    Dim output_file_path As String
    Dim output_sheet As Excel.Worksheet
    Dim output_workbook As Excel.Workbook
    Dim safety_counter As Integer
    Dim source_excel As Scripting.Dictionary
    Dim source_workbook As Excel.Workbook
    Dim source_sheet As Excel.Worksheet
    Dim source_sheet_name As String
    Dim str_file_name As String
    Dim str_file_name_premium As String
    Dim str_report_month As String
    Dim str_report_year As String
    Dim output_data() As Variant
    Dim row_counter As Long
    Dim unique_currencies As New Collection
        
    awac_v5_name_test = "AW RiskPoint Premium"
    awac_account_current_name_test = "AW RiskPoint Account Current"
    awac_payment_name_test = "AW Payment Report"
    
    'Loop through all files in the folder
    folder_path = Load.system_info.system_paths.awac_conversion_path
    filename = Dir(folder_path & "*.xlsx")
    safety_counter = 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set Me.col_file_list_for_email_attaching = Nothing
    Set app_excel = New Excel.Application
    Set col_source_files = Nothing
    Do While filename <> "" And safety_counter < 20
        safety_counter = safety_counter + 1
        
        If Not source_sheet Is Nothing Then Set source_sheet = Nothing
        Set source_workbook = app_excel.Workbooks.Open(folder_path & filename)
        
        'check if workbook is written or payment. If neither, go to end of loop
        Me.current_source_report_type = "-1"
        On Error Resume Next
            Set source_sheet = source_workbook.Sheets(v5_source_sheet)
            Set source_sheet = source_workbook.Sheets(v5_source_sheet_old)
            Set source_sheet = source_workbook.Sheets(payment_source_sheet)
            Set source_sheet = source_workbook.Sheets(payment_source_sheet_old)
        On Error GoTo err_handler
        If Load.is_debugging = True Then On Error GoTo 0
        
        'check if a sheet was found
        If source_sheet Is Nothing Then
            GoTo end_of_loop
        End If
        
        source_sheet_name = source_sheet.Name
        last_row = source_sheet.Cells(source_sheet.Rows.Count, "A").End(xlUp).Row
        excel_data = source_sheet.Range("A1:DA" & last_row)
        
        'reporting period is in B3, in this format: 01-Jul-2024 to 31-Jul-2024
        str_report_month = awac_excels.convert_month_name_to_xx(Split(Split(excel_data(3, 2), "to")(0), "-")(1))
        str_report_year = (Split(Split(excel_data(3, 2), " to")(0), "-")(2))
        
        Me.found_v5_report = False
        Me.found_payment_report = False
        If source_sheet_name = v5_source_sheet Or source_sheet_name = v5_source_sheet_old Then
            currency_column = 30
            current_source_report_type = Me.reports.written_report
            Me.found_v5_report = True
            Me.reports.v5_original_file_path = folder_path & filename
            Me.reports.v5_original_file_name = filename
        ElseIf source_sheet_name = payment_source_sheet Or source_sheet_name = payment_source_sheet_old Then
            currency_column = 21
            current_source_report_type = Me.reports.payment_report
            Me.found_payment_report = True
            Me.reports.payment_original_file_path = folder_path & filename
            Me.reports.payment_original_file_name = filename
        End If
        
        source_workbook.Close
        Set source_workbook = Nothing
        
        current_binder = Me.emails_convert_excels_binder_identification(excel_data, is_test_run)
        
        If current_binder = "-1" Then
            MsgBox "I think " & filename & " is a valid report, but I could not figure out which binder it is for." _
            & vbNewLine & vbNewLine & "source: " & proc_name _
            , , "Could not identify binder"
            GoTo end_of_loop
        End If
        
        'making a list of currencies
        Set unique_currencies = Nothing
        With unique_currencies
            For row_counter = 12 To last_row
                On Error Resume Next
                .Add excel_data(row_counter, currency_column), CStr(excel_data(row_counter, currency_column))
                On Error GoTo err_handler
                If Load.is_debugging = True Then On Error GoTo 0
            Next row_counter
        End With
        
        '15 Jan 2025, CK: _
            - Our one V5 Writtenr report shall be broken into one report per currency for Canada. _
            - The report for USA shall ever only contain one currency, USD. _
            - The file name of the report(s) for Canada shall contain the currency. _
            - The file name of the US report shall not contain the currency.
            
        If Me.current_source_report_type = Me.reports.written_report Then
            file_name_component = "Account Current"
            file_name_component_2 = " RiskPoint"
        ElseIf Me.current_source_report_type = Me.reports.payment_report Then
            file_name_component = "Payment Report"
            file_name_component_2 = ""
        End If
        
        For Each binder_currency In unique_currencies
            output_data = Me.emails_convert_excels_prepare_output_table(excel_data _
            , output_data _
            , binder_currency _
            , Me.current_source_report_type _
            , currency_column)
            
            '15 Jan 2025, CK: The AW RiskPoint Premium report shall only be generated if Me.current_source_report_type = Me.reports.written_report
            
            If current_binder = global_vars.awac_excels_binders.awac_can Then
                str_file_name = "AW" & file_name_component_2 & " " & binder_currency & " " & file_name_component & " -" & str_report_month & "-" & str_report_year & ".xlsx"
                str_file_name_premium = "AW" & file_name_component_2 & " " & binder_currency & " Premium" & " -" & str_report_month & "-" & str_report_year & ".xlsx"
            ElseIf current_binder = global_vars.awac_excels_binders.awac_usa Then
                str_file_name = "AW" & file_name_component_2 & " " & file_name_component & " -" & str_report_month & "-" & str_report_year & ".xlsx"
                str_file_name = Replace(str_file_name, "  ", " ")
                str_file_name_premium = "AW " & file_name_component_2 & " Premium" & " -" & str_report_month & "-" & str_report_year & ".xlsx"
                str_file_name_premium = Replace(str_file_name_premium, "  ", " ")
            End If
            
            output_file_path = Me.current_output_folder & str_file_name
            Me.emails_convert_excels_prepare_converted_file output_data, output_file_path, app_excel
            
            Set excel_file_for_email = New Scripting.Dictionary
            With excel_file_for_email
                .Add "file_name", str_file_name
                .Add "folder_path", Me.current_output_folder
            End With
            With col_file_list_for_email_attaching
                .Add excel_file_for_email
            End With
            
            If Me.current_source_report_type = Me.reports.written_report Then
                output_file_path = Me.current_output_folder & str_file_name_premium
                
                'need code to create RP V5 report tables per currency
                '17 Jan 2025, CK: consider passing output_data by reference and turning it into a sub
                output_data = Me.emails_convert_excels_prepare_output_table_v5(excel_data, binder_currency, currency_column)
                Me.emails_convert_excels_prepare_converted_file_v5 output_data, output_file_path, app_excel
                
                'add file to list for later attaching to emails
                Set excel_file_for_email = New Scripting.Dictionary
                With excel_file_for_email
                    .Add "file_name", str_file_name_premium
                    .Add "folder_path", Me.current_output_folder
                End With
                With Me.col_file_list_for_email_attaching
                    .Add excel_file_for_email
                End With
            End If
            
        Next binder_currency
        
        Me.error_excel_conversion = False
        
        'add source file to list for mvoing to archive
        Set source_excel = New Scripting.Dictionary
        With source_excel
            .Add "target_folder_path", Me.current_source_folder
            .Add "source_file_path", folder_path & filename
            .Add "file_name", filename
        End With
        With col_source_files
            .Add source_excel
        End With
end_of_loop:
        filename = Dir 'Move to the next file in the folder
    Loop
    
    'move source files to archive
    '17 Jan 2025, CK: The dir command will fail if files are moved while the loop is running. Hence, instead of copying during the loop, I'll store _
    all files in a collection with dicts, to loop through and move when the loop is done.
    Dim file_name_addition As String
    Dim source_excel_file_name As String
    For Each source_excel In col_source_files
        source_excel_file_name = source_excel("file_name")
        source_excel_file_name = Split(source_excel_file_name, ".")(0)
        
        file_name_addition = ""
        i = 1
        file_name = source_excel_file_name & file_name_addition & ".xlsx"
        Do Until Dir(source_excel("target_folder_path") & file_name) = ""
            
            file_name_addition = "_" & i
            i = i + 1
            file_name = source_excel_file_name & file_name_addition & ".xlsx"
        Loop
        str_milestone = "Trying to move source file."
        file_mover.MoveFile source_excel("source_file_path"), source_excel("target_folder_path") & file_name
        
    Next source_excel
    
outro:
    If Not app_excel Is Nothing Then Set app_excel = Nothing
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .call_stack = Load.call_stack
        .routine_name = proc_name
        .milestone = str_milestone
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    Me.error_excel_conversion = True
    GoTo outro
End Sub
Public Sub emails_convert_excels_prepare_converted_file(ByVal output_data _
, ByVal output_file_path As String _
, ByVal app_excel As Excel.Application)
    
    Dim output_workbook As Excel.Workbook
    Dim output_sheet As Excel.Worksheet
    
    Set output_workbook = app_excel.Workbooks.Add
    output_workbook.Activate
    Set output_sheet = app_excel.Worksheets(1)
    
    With output_sheet
        .Columns(1).ColumnWidth = 20
        .Columns(2).ColumnWidth = 35
        .Columns(3).ColumnWidth = 11
        .Columns(4).ColumnWidth = 10
        .Columns(5).ColumnWidth = 15
        .Columns(6).ColumnWidth = 13
        .Columns(7).ColumnWidth = 13
        .Columns(8).ColumnWidth = 13
        .Columns(9).ColumnWidth = 13
        .Columns(10).ColumnWidth = 13
        .Columns(11).ColumnWidth = 13
        .Columns(12).ColumnWidth = 13
        .Columns(13).ColumnWidth = 13
        .Columns(14).ColumnWidth = 13
        .Columns(15).ColumnWidth = 13
        .Columns(16).ColumnWidth = 20
        .Columns(17).ColumnWidth = 26
        .Columns(18).ColumnWidth = 13
        .Columns(19).ColumnWidth = 13
        .Columns(19).ColumnWidth = 13
        .Range(.Cells(1, 1), .Cells(UBound(output_data), 19)) = output_data
    End With
    With output_workbook
        .SaveAs output_file_path, 51
        .Close
    End With
        
End Sub
Public Sub emails_convert_excels_prepare_converted_file_v5(ByVal output_data _
, ByVal output_file_path As String _
, ByVal app_excel As Excel.Application)
    
    '27 Jan 2025, CK: CK TO FIX ON ERROR, beacuse if file is in use, it will crash
    Dim output_workbook As Excel.Workbook
    Dim output_sheet As Excel.Worksheet
    Dim last_row As Long
    Dim last_column As Integer
    last_row = UBound(output_data)
    last_column = UBound(output_data, 2)
    Set output_workbook = app_excel.Workbooks.Add
    output_workbook.Activate
    Set output_sheet = app_excel.Worksheets(1)
    
    With output_sheet
        .Range(.Cells(1, 1) _
        , .Cells(last_row, last_column) _
        ) = output_data
    End With
    With output_workbook
        .SaveAs output_file_path, 51
        .Close
    End With
        
End Sub
Public Function emails_convert_excels_prepare_output_table_v5(ByVal excel_data _
, ByVal input_currency As String _
, ByVal currency_column As Integer)
    Dim proc_name As String
    proc_name = "cls_form_awac_excels.emails_convert_excels_prepare_output_table_v5"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim column_counter As Long
    Dim excel_row_counter As Long
    Dim last_row As Long
    Dim last_column As Integer
    Dim row_counter As Long
    
    excel_row_counter = 2
    last_row = UBound(excel_data)
    last_column = UBound(excel_data, 2)
    ReDim output_data(1 To last_row - 10, 1 To last_column)
        
    'add headings to output
    For column_counter = 1 To last_column
        output_data(1, column_counter) = excel_data(11, column_counter)
    Next column_counter
    
    'add data
    excel_row_counter = 2
    For row_counter = 2 To last_row - 10
        If excel_data(row_counter + 10, currency_column) = input_currency Then
            For column_counter = 1 To last_column
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, column_counter)
            Next column_counter
            excel_row_counter = excel_row_counter + 1
        End If
        column_counter = 1
    Next row_counter
    
    emails_convert_excels_prepare_output_table_v5 = output_data
    
outro:
    Exit Function

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Function
Public Function emails_convert_excels_prepare_output_table(ByVal excel_data _
, ByVal output_data _
, ByVal input_currency As String _
, ByVal source_report_type As String _
, ByVal currency_column)
    Dim proc_name As String
    proc_name = "cls_form_awac_excels.emails_convert_excels_prepare_output_table"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim column_counter As Long
    Dim excel_row_counter As Long
    Dim last_row As Long
    Dim row_counter As Long
    
    excel_row_counter = 2
    last_row = UBound(excel_data)
    
    ReDim output_data(1 To last_row, 1 To 20)
        
    'add headings to output
    column_counter = 1
    output_data(1, column_counter) = "Policy Number": column_counter = column_counter + 1
    output_data(1, column_counter) = "Insured Name": column_counter = column_counter + 1
    output_data(1, column_counter) = "Billing Date": column_counter = column_counter + 1
    output_data(1, column_counter) = "Invoice #": column_counter = column_counter + 1
    output_data(1, column_counter) = "Tran. Type": column_counter = column_counter + 1
    output_data(1, column_counter) = "Effective Date": column_counter = column_counter + 1
    output_data(1, column_counter) = "Expiration Date": column_counter = column_counter + 1
    output_data(1, column_counter) = "Total Premium": column_counter = column_counter + 1
    output_data(1, column_counter) = "Program Comm %": column_counter = column_counter + 1
    output_data(1, column_counter) = "Comm Amt": column_counter = column_counter + 1
    output_data(1, column_counter) = "State Surcharges": column_counter = column_counter + 1
    output_data(1, column_counter) = "Misc Taxes": column_counter = column_counter + 1
    output_data(1, column_counter) = "Policy Fee": column_counter = column_counter + 1
    output_data(1, column_counter) = "Misc Fee": column_counter = column_counter + 1
    output_data(1, column_counter) = "Net Premium": column_counter = column_counter + 1
    output_data(1, column_counter) = "Product": column_counter = column_counter + 1
    output_data(1, column_counter) = "Carrier": column_counter = column_counter + 1
    output_data(1, column_counter) = "Payment Plan": column_counter = column_counter + 1
    output_data(1, column_counter) = "Due Date": column_counter = column_counter + 1
    output_data(1, column_counter) = "National Producer Number (NPN)": column_counter = column_counter + 1
    
    Dim closing_date_column As Integer
    closing_date_column = 77
    If source_report_type = Me.reports.written_report Then
        For row_counter = 2 To last_row - 10
            column_counter = 1
    
            'check if closing date is populated. If not, ignore row.
            If excel_data(row_counter + 10, closing_date_column) <> "" _
            And excel_data(row_counter + 10, currency_column) = input_currency Then
            
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 20)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 25)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 18)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 22)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 23)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 41)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 49)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 48)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 50) 'CR1280
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 15)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 5)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "Full Payment"
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                excel_row_counter = excel_row_counter + 1
            End If
        Next row_counter
    ElseIf source_report_type = Me.reports.payment_report Then
        For row_counter = 2 To last_row - 10
            If excel_data(row_counter + 10, currency_column) = input_currency Then
                
                column_counter = 1
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 25)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 30)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 23)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 27)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 28)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 22)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 54)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 55)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 65)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 18)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = excel_data(row_counter + 10, 5)
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "Full Payment"
                column_counter = column_counter + 1
                
                output_data(excel_row_counter, column_counter) = "" ' intentionally blank
                
                excel_row_counter = excel_row_counter + 1
            End If
        Next row_counter
    End If
    
    emails_convert_excels_prepare_output_table = output_data
    
outro:
    Exit Function

err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = proc_name
        .milestone = ""
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Function
Public Function convert_month_name_to_xx(ByVal month_name As String) As String
    Dim output As String
    If month_name = "jan" Then output = "01"
    If month_name = "feb" Then output = "02"
    If month_name = "mar" Then output = "03"
    If month_name = "apr" Then output = "04"
    If month_name = "may" Then output = "05"
    If month_name = "jun" Then output = "06"
    If month_name = "jul" Then output = "07"
    If month_name = "aug" Then output = "08"
    If month_name = "sep" Then output = "09"
    If month_name = "oct" Then output = "10"
    If month_name = "nov" Then output = "11"
    If month_name = "dec" Then output = "12"
    convert_month_name_to_xx = output
End Function
Public Sub sort_all_fields_in_form_for_debugging()
    'to sort the fields, set the form to design mode, and run this module
    
    Dim form_field As cls_field
    Load.check_conn_and_variables
    Me.init
    Me.make_all_fields_visible_for_debugging = True
    Load.is_debugging = True
    For Each form_field In Me.col_all_fields
        With form_field
            .field_value = True
        End With
    Next form_field
    With Me.header_instructions
        .field_left = 10000
    End With
    With Me.header_cmd_goto_folder
        .field_left = Me.header_instructions.field_right
    End With
    Me.paint
    Me.make_all_fields_visible_for_debugging = False
End Sub
Public Sub emails()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_forms_awac_excels.emails"
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Me.emails_prepare Me.reports.payment_report
    Me.emails_prepare Me.reports.written_report
    
outro:
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .call_stack = Load.call_stack
        .routine_name = "cls_forms_awac_excels.emails"
        .milestone = ""
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
End Sub

Public Function emails_prepare(ByVal report_type As String) As String
    Load.call_stack = Load.call_stack & vbNewLine & "cls_forms_awac_excels.emails_prepare"
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_to As String, str_cc As String, str_sql As String, rs As ADODB.Recordset, str_subject As String, str_body As String
    If report_type = Me.reports.payment_report Then
        str_sql = "SELECT awac_id, email_address, to_field_id__menu_list_t, cc_field_id__menu_list_t" _
            & " FROM " & Load.sources.awac_excels_table _
            & " WHERE is_deleted = 0 AND payment_report_id__menu_list_t =  " & menu_list.yes
        str_subject = "RP | AWAC: Payment Report"
        str_body = "<p>Hi, </p><p>Please find attached the payment report from RiskPoint.</p><p>Best regards</p>"
    ElseIf report_type = Me.reports.written_report Then
        str_sql = "SELECT awac_id, email_address, to_field_id__menu_list_t, cc_field_id__menu_list_t" _
            & " FROM " & Load.sources.awac_excels_table _
            & " WHERE is_deleted = 0 AND v5_report_id__menu_list_t =  " & menu_list.yes
        str_subject = "RP | AWAC: Account Current Reports"
        str_body = "<p>Hi, </p><p>Please find attached the account current reports from RiskPoint.</p><p>Best regards</p>"
    Else
        report_type = "-1"
        GoTo outro
    End If
    
    Set rs = utilities.create_adodb_rs(conn, str_sql)
    With rs
        str_to = ""
        str_cc = ""
        Do Until .EOF
            If !to_field_id__menu_list_t = menu_list.yes Then str_to = str_to & !email_address & "; "
            If !cc_field_id__menu_list_t = menu_list.yes Then str_cc = str_cc & !email_address & "; "
            .MoveNext
        Loop
        .Close
    End With
    
    'Create email object
    Dim Outlook As Object, email As Object
    Set Outlook = CreateObject("Outlook.Application")
    Set email = Outlook.CreateItem(0)
    With email
        .To = str_to
        .CC = str_cc
        .BCC = "-1"
        .Subject = str_subject
        .HTMLBody = str_body
        .Display
    End With
    
    Set Outlook = Nothing
    Set email = Nothing

outro:
    Exit Function
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .call_stack = Load.call_stack
        .routine_name = "cls_forms_awac_excels.emails_prepare"
        .milestone = ""
        .params = "str_sql = " & str_sql
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    
End Function