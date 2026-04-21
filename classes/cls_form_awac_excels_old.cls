Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public form_name As String
Public make_all_fields_visible_for_debugging As Boolean
Public reports As New cls_form_awac_excels_reports
Public error_excel_conversion As Boolean
Public found_payment_report As Boolean
Public found_v5_report As Boolean

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
Public header_cmd_show_input As New cls_field

Public header_instructions As New cls_field

Public col_input_fields As New Collection
Public col_all_fields As New Collection
Public col_cmds As New Collection

Public folder_path_awac_canada_source As String
Public folder_path_awac_canada_output As String
Public folder_path_awac_usa_source As String
Public folder_path_awac_usa_output As String

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

Public Sub init_excel_variables()
    folder_path_awac_canada_source = Load.system_info.system_paths.awac_conversion_path & "awac_can__archive\reports_source\"
    folder_path_awac_canada_output = Load.system_info.system_paths.awac_conversion_path & "awac_can__archive\reports_output\"
    folder_path_awac_usa_source = Load.system_info.system_paths.awac_conversion_path & "awac_usa__archive\reports_source\"
    folder_path_awac_usa_output = Load.system_info.system_paths.awac_conversion_path & "awac_usa__archive\reports_output\"
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
        .field_name = "header_instructions"
        .field_visible = True
        .field_width = 4500
        .field_top = 100
        .field_left = 100
    End With
    With Me.header_cmd_goto_folder
        .field_name = "cmd_goto_folder"
        .field_visible = True
        .field_left = Me.header_instructions.field_right
        .field_top = Me.header_instructions.field_top
    End With
    
    With Me.header_cmd_add_change
        .field_name = "cmd_add_change"
        .field_visible = False
        .field_top = 100
        .field_left = 5000
    End With
    With Me.header_cmd_show_input
        .field_name = "cmd_show_input"
        .field_visible = True
        .field_top = 1000
        .field_left = 8000
    End With
    With Me.header_cmd_emails
        .field_name = "cmd_prepare_emails"
        .field_visible = True
        .field_top = 100
        .field_left = 8000
    End With
    
    Set form_field = Nothing
End Sub

Public Sub paint()
    Load.call_stack = Load.call_stack & vbNewLine & "cls_forms_awac_excels.paint"
    Dim form_field As cls_field
    
    For Each form_field In Me.col_all_fields
        With Forms(Me.form_name).Controls(form_field.field_name)
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
Public Sub emails_convert_excels(ByVal is_test_run As Boolean)
    Load.call_stack = Load.call_stack & vbNewLine & "cls_forms_awac_excels.convert_excels"
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    'there are two source excels. The code will identify them by their sheet names
    Dim v5_source_sheet As String, payment_source_sheet As String
    v5_source_sheet = "V5WrittenReport"
    payment_source_sheet = "V5PayableReport"
    Me.found_v5_report = False
    Me.found_payment_report = False
    
    Dim source_workbook As Excel.Workbook, source_sheet As Excel.Worksheet, folder_path As String
    Dim output_workbook As Excel.Workbook, output_sheet As Excel.Worksheet, output_file_path As String
    Dim last_row As Long, str_report_month As String, str_report_year As String
    Dim excel_data() As Variant, output_data() As Variant
    Dim row_counter As Long, column_counter As Long
    Dim awac_v5_name_test As String, awac_account_current_name_test As String, awac_payment_name_test As String
    awac_v5_name_test = "AW RiskPoint Premium"
    awac_account_current_name_test = "AW RiskPoint Account Current"
    awac_payment_name_test = "AW Payment Report"
    folder_path = Load.system_info.system_paths.awac_conversion_path
    
    'Get the first file in the folder
    Dim filename As String
    filename = Dir(folder_path & "*.xlsx")
    
    Dim app_excel As Excel.Application
    Set app_excel = New Excel.Application
    
    'fso is used for copying files later in the code
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Loop through all files in the folder
    Dim safety_counter As Integer, excel_row_counter As Long
    safety_counter = 0
    Do While filename <> "" And safety_counter < 10
        safety_counter = safety_counter + 1
        
        If InStr(0, awac_v5_name_test) = 0 And InStr(0, awac_account_current_name_test) = 0 And InStr(0, awac_payment_name_test) = 0 Then
            If Not source_sheet Is Nothing Then Set source_sheet = Nothing
            Set source_workbook = app_excel.Workbooks.Open(folder_path & filename)
            
            'check if workbook is written or payment. If neither, go to end of loop
            On Error Resume Next
                Set source_sheet = source_workbook.Sheets(v5_source_sheet)
                Set source_sheet = source_workbook.Sheets(payment_source_sheet)
            On Error GoTo err_handler
            If Load.is_debugging = True Then On Error GoTo 0
            
            'check if a sheet was found
            If source_sheet Is Nothing Then
                GoTo end_of_loop
            End If
            
            last_row = source_sheet.Cells(source_sheet.Rows.Count, "A").End(xlUp).Row
            excel_data = source_sheet.Range("A1:DA" & last_row)
            Dim source_sheet_name As String
            source_sheet_name = source_sheet.name
            source_workbook.Close
            Set source_workbook = Nothing
            
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
            
            If source_sheet_name = v5_source_sheet Then
                Me.found_v5_report = True
                Me.reports.v5_original_file_path = folder_path & filename
                Me.reports.v5_original_file_name = filename
                'make a copy of source file and name it according to awac's naming convention.
                'reporting period is in B3, in this format: 01-Jul-2024 to 31-Jul-2024
                str_report_month = awac_excels.convert_month_name_to_xx(Split(Split(excel_data(3, 2), "to")(0), "-")(1))
                str_report_year = (Split(Split(excel_data(3, 2), " to")(0), "-")(2))
                
                'copy file with awac naming convention and top 10 rows removed t archive
                output_file_path = folder_path & "archive\AW RiskPoint Premium " & str_report_month & "-" & str_report_year & ".xlsx"
                fso.CopyFile folder_path & filename, output_file_path
                Me.emails_remove_top10_rows_from_v5_written output_file_path, app_excel
                Me.reports.v5_renamed_file_path = output_file_path
                
                'store path of converted file for later attaching to email
                '10 Jan 2025, CK: Need to implement logic to cater for different paths for different binders
                output_file_path = folder_path & "archive\AW RiskPoint Account Current -" & str_report_month & "-" & str_report_year & ".xlsx"
                Me.reports.v5_converted_file_path = output_file_path
                
                output_data = Me.emails_convert_excels_written(excel_data, output_data)
                
            End If
            
            If source_sheet_name = payment_source_sheet Then
                Me.found_payment_report = True
                
                For row_counter = 2 To last_row - 10
                    column_counter = 1
                    output_data(row_counter, column_counter) = excel_data(row_counter + 10, 25)
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = excel_data(row_counter + 10, 30)
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = "" ' intentionally blank
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = "" ' intentionally blank
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = excel_data(row_counter + 10, 23)
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = excel_data(row_counter + 10, 27)
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = excel_data(row_counter + 10, 28)
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = excel_data(row_counter + 10, 22)
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = excel_data(row_counter + 10, 54)
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = excel_data(row_counter + 10, 55)
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = "" ' intentionally blank
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = "" ' intentionally blank
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = "" ' intentionally blank
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = "" ' intentionally blank
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = excel_data(row_counter + 10, 65)
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = excel_data(row_counter + 10, 18)
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = excel_data(row_counter + 10, 5)
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = "Full Payment"
                    column_counter = column_counter + 1
                    
                    output_data(row_counter, column_counter) = "" ' intentionally blank
                Next row_counter
            End If
            
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
                .Range(.Cells(1, 1), .Cells(last_row, 19)) = output_data
            End With
            
            'reporting period is in A3, in this format: '16-Jul-2024 to 31-Jul-2024'
            str_report_month = awac_excels.convert_month_name_to_xx(Split(Split(excel_data(3, 2), "to")(0), "-")(1))
            str_report_year = (Split(Split(excel_data(3, 2), " to")(0), "-")(2))
            output_file_path = folder_path & "archive\AW Payment Report " & str_report_month & "-" & str_report_year & ".xlsx"
                
            With output_workbook
                .SaveAs output_file_path, 51
                .Close
            End With
            
            
            Me.reports.payment_original_file_path = folder_path & filename
                Me.reports.payment_original_file_name = filename
                
                
                Me.reports.payment_converted_file_path = output_file_path
end_of_loop:
            filename = Dir 'Move to the next file in the folder

        End If
    Loop
    
    'move source files to archive
    Dim file_mover As New FileSystemObject, file_name_addition As String, i As Integer, file_name As String
    If Me.found_v5_report = True Then
        file_name_addition = ""
        i = 1
        file_name = Split(Me.reports.v5_original_file_name, ".xlsx")(0) & file_name_addition & ".xlsx"
        Do Until Dir(folder_path & "archive\" & file_name) = ""
            file_name_addition = "_" & i
            i = i + 1
            file_name = Split(Me.reports.v5_original_file_name, ".xlsx")(0) & file_name_addition & ".xlsx"
        Loop
        file_mover.MoveFile Me.reports.v5_original_file_path, folder_path & "archive\" & file_name
    End If
    
    If Me.found_payment_report = True Then
        file_name_addition = ""
        i = 1
        file_name = Split(Me.reports.payment_original_file_name, ".xlsx")(0) & file_name_addition & ".xlsx"
        Do Until Dir(folder_path & "archive\" & file_name) = ""
            file_name_addition = "_" & i
            i = i + 1
            file_name = Split(Me.reports.payment_original_file_name, ".xlsx")(0) & file_name_addition & ".xlsx"
        Loop
        file_mover.MoveFile Me.reports.payment_original_file_path, folder_path & "archive\" & file_name
    End If
    Me.error_excel_conversion = False
    
outro:
    If Not app_excel Is Nothing Then Set app_excel = Nothing
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .call_stack = Load.call_stack
        .routine_name = "cls_forms_awac_excels.convert_excels"
        .milestone = ""
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    Me.error_excel_conversion = True
    GoTo outro
End Sub
Public Function emails_convert_excels_written(ByVal excel_data, ByVal output_data)
    Dim proc_name As String
    proc_name = "cls_form_awac_excels.emails_convert_excels_written"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
        
    Dim excel_row_counter As Long, row_counter As Long, last_row As Long, column_counter As Long
    excel_row_counter = 2
    For row_counter = 2 To last_row - 10
        column_counter = 1

        'check if closing date is populated. If not, ignore row.
        If excel_data(row_counter + 10, 77) <> "" Then
        
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
Public Sub emails_remove_top10_rows_from_v5_written(ByVal file_path As String, ByVal app_excel As Excel.Application)
    Load.call_stack = Load.call_stack & vbNewLine & "cls_forms_awac_excels.emails_remove_top10_rows_from_v5_written"
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim source_workbook As Excel.Workbook, source_sheet As Excel.Worksheet
    Set source_workbook = app_excel.Workbooks.Open(file_path)
    Set source_sheet = source_workbook.Worksheets(1)
    With source_sheet
        .Rows("1:10").Delete
    End With
    With source_workbook
        .Close True
    End With
    Set source_sheet = Nothing
    Set source_workbook = Nothing
    
outro:
    If Not app_excel Is Nothing Then Set app_excel = Nothing
    Exit Sub
    
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .call_stack = Load.call_stack
        .routine_name = "cls_forms_awac_excels.emails_remove_top10_rows_from_v5_written"
        .milestone = ""
        .params = "file_path = " & file_path
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    Me.error_excel_conversion = True
    GoTo outro
    
End Sub
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
    
    'Me.find_output_folders
        'binder parent folder
        'reports_source
        'reports_output
    
    Me.emails_convert_excels False
    If Me.error_excel_conversion = True Then
        GoTo outro
    End If
    
    If Me.found_payment_report = True Then Me.emails_prepare Me.reports.payment_report, Me.reports.payment_converted_file_path
    If Me.found_v5_report = True Then Me.emails_prepare Me.reports.written_report, Me.reports.v5_converted_file_path, Me.reports.v5_renamed_file_path
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

Public Function emails_prepare(ByVal report_type As String, ByVal attachment_1_path As String, Optional ByVal attachment_2_path As String) As String
    Load.call_stack = Load.call_stack & vbNewLine & "cls_forms_awac_excels.emails_prepare"
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    'check if paths are valid
    If Dir(attachment_1_path) = "" Then
        emails_prepare = "-1"
        GoTo outro
    End If
    If attachment_2_path <> "" Then
        If Dir(attachment_2_path) = "" Then
            emails_prepare = "-1"
            GoTo outro
        End If
    End If
    
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
        .Attachments.Add attachment_1_path
        If attachment_2_path <> "" Then .Attachments.Add attachment_2_path
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