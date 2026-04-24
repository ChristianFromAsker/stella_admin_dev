Option Compare Database
Option Explicit

'used for get_event_id
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef GUID As GUID) As LongPtr
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Type typ_log_object
    change_source As String
    changer_id As String
    comment As String
    data_set As String
    deal_id As Long
    event_id As String
    executed_sql As String
    field_name As String
    new_value As Variant
    operation_type As String
    policy_id As Long
    record_id As Long
    security_id As Long
End Type
Function IsWithinFivePercent(ByVal x As Double, ByVal y As Double, _
                             Optional ByVal Epsilon As Double = 0.000001) As Boolean
    If y = 0 Then
        IsWithinFivePercent = (Abs(x) <= Epsilon)
    Else
        IsWithinFivePercent = (Abs(x - y) <= 0.05 * Abs(y))
    End If
End Function
Public Function levenshtein(ByVal s1 As String _
    , ByVal s2 As String _
    , Optional ByVal max_distance As Long _
) As Long
    Const proc_name As String = "utilities.levenshtein"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim i As Long
    Dim j As Long
    Dim output As Long
    Dim s1_len As Long
    Dim s2_len As Long
    Dim cost As Long
    Dim d() As Long
    
    s1_len = Len(s1)
    s2_len = Len(s2)
    
    If s1_len = 0 Then
        output = s2_len
        GoTo outro
    End If
    If s2_len = 0 Then
        output = s1_len
        GoTo outro
    End If
    If max_distance <> 0 Then
        If s1_len - s2_len > max_distance Or s2_len - s1_len > max_distance Then
            output = max_distance + 1
            GoTo outro
        End If
    End If
    
    ReDim d(0 To s1_len, 0 To s2_len)
    
    For i = 0 To s1_len
        d(i, 0) = i
    Next i
    
    For j = 0 To s2_len
        d(0, j) = j
    Next j
    
    For i = 1 To s1_len
        For j = 1 To s2_len
            If Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            
            d(i, j) = d(i - 1, j) + 1               ' deletion
            If d(i, j) > d(i, j - 1) + 1 Then d(i, j) = d(i, j - 1) + 1   ' insertion
            If d(i, j) > d(i - 1, j - 1) + cost Then d(i, j) = d(i - 1, j - 1) + cost ' substitution
            
            'exit logic if distance
            If max_distance <> 0 Then
                If d(s1_len, s2_len) > max_distance Then
                    output = max_distance + 1
                    GoTo outro
                End If
            End If
        Next j
    Next i
    
    output = d(s1_len, s2_len)
outro:
    levenshtein = output
    utilities.call_stack_remove_last_item
    Exit Function
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    output = -1
    Resume outro
End Function
Public Function simplify_company_name(ByVal input_name As String) As String
    Const proc_name As String = "utilities.simplify_company_name"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim col_company_endings As New Collection
    Set col_company_endings = Nothing
    With col_company_endings
        .Add "(pty)"
        .Add "AB"
        .Add "ApS"
        .Add "AS"
        .Add "ASA"
        .Add "CCorp"
        .Add "CIC"
        .Add "Co Inc"
        .Add "Co Ltd"
        .Add "Corp"
        .Add "GP"
        .Add "GmbH"
        .Add "Inc"
        .Add "JSC"
        .Add "KK"
        .Add "KG"
        .Add "Lda"
        .Add "LLC"
        .Add "LLP"
        .Add "LP"
        .Add "Ltd"
        .Add "NL"
        .Add "OJSC"
        .Add "Pty Ltd"
        .Add "Pvt Ltd"
        .Add "SA"
        .Add "sarl"
        .Add "SAS"
        .Add "SE"
        .Add "SPV"
        .Add "SNC"
        .Add "SRA"
    End With
    
    Dim ending_check As String
    Dim output_name As String
    output_name = Replace(input_name, ".", "")
    output_name = Replace(output_name, ",", "")
    output_name = Replace(output_name, "'", "")
    output_name = Replace(output_name, ":", "")
    output_name = Replace(output_name, ";", "")
    output_name = Replace(output_name, "/", "")
    output_name = Trim(output_name)
    
    Dim company_ending As Variant
    Dim found_company_ending As String
    Dim found_company_ending_length As Integer
    
    found_company_ending_length = 0
    For Each company_ending In col_company_endings
        If Right(output_name, Len(" " & company_ending)) = " " & company_ending Then
            If Len(" " & company_ending) > found_company_ending_length Then
                found_company_ending_length = Len(" " & company_ending)
                found_company_ending = " " & company_ending
            End If
        End If
    Next company_ending
    
    output_name = Left(output_name, Len(output_name) - Len(found_company_ending))
    simplify_company_name = output_name
    
outro:
    utilities.call_stack_remove_last_item
    Exit Function
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Function
Public Function clean_entity_name(ByVal str_input As String) As String
    Const proc_name As String = "utilities.clean_entity_name"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim allowed_chars As String
    Dim ch_found As Boolean
    Dim col_allowed_chars As Collection
    Dim i As Long
    Dim obj_ch As Variant
    Dim rs As ADODB.Recordset
    Dim str_allowed_chars As String
    Dim str_ch As String
    Dim str_output As String
    Dim str_sql As String
    Dim test As String
    
    'IMPORTANT NOTE FROM CK! This is used in stella_uw and in cm_admin for uploading lists of bad entities. _
    It is imperativ that the lookup and upload function use the same cleaner.
    
    allowed_chars = utilities.get_txt_from_file(Load.system_info.system_paths.stella_path & "settings\allowed_chars.txt")
    
    Set col_allowed_chars = New Collection
    For i = 1 To Len(allowed_chars)
        str_ch = Mid(allowed_chars, i, 1)
        col_allowed_chars.Add Asc(str_ch)
    Next i
    
    str_output = ""
    For i = 1 To Len(str_input)
        str_ch = Mid(str_input, i, 1)
        ch_found = False
        For Each obj_ch In col_allowed_chars
            If Asc(str_ch) = obj_ch Then
                ch_found = True
            End If
        Next obj_ch
        If ch_found = True Then
            str_output = str_output & str_ch
        End If
    Next i
    
    Do While InStr(str_output, "  ") > 0
        str_output = Replace(str_output, "  ", " ")
    Loop
    
    clean_entity_name = str_output
outro:
    utilities.call_stack_remove_last_item False
    Exit Function
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Function
Public Function convert_layer_to_words(ByVal int_layer As Integer) As String
    Const proc_name As String = "utilities.convert_layer_to_words"
    utilities.call_stack_add_item proc_name
    
    Dim str_return As String
    If int_layer = 0 Then
        str_return = "Primary"
    ElseIf int_layer = 1 Then
        str_return = "1st xs"
    ElseIf int_layer = 2 Then
        str_return = "2nd xs"
    ElseIf int_layer = 3 Then
        str_return = "3rd xs"
    ElseIf int_layer = 4 Then
        str_return = "4th xs"
    ElseIf int_layer = 5 Then
        str_return = "5th xs"
    ElseIf int_layer = 6 Then
        str_return = "6th xs"
    ElseIf int_layer = 7 Then
        str_return = "7th xs"
    ElseIf int_layer = 8 Then
        str_return = "8th xs"
    ElseIf int_layer = 9 Then
        str_return = "9th xs"
    ElseIf int_layer = 10 Then
        str_return = "10th xs"
    ElseIf int_layer = 11 Then
        str_return = "11th xs"
    ElseIf int_layer = 12 Then
        str_return = "12th xs"
    End If
    convert_layer_to_words = str_return
outro:
    utilities.call_stack_remove_last_item False
    Exit Function
End Function

Public Function get_txt_from_file(ByVal file_path As String) As String
    Const proc_name As String = "utilities.get_txt_from_file"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim output As String
    Dim stm As ADODB.Stream
    
    Set stm = New ADODB.Stream
    With stm
        .Type = 2            ' adTypeText
        .Charset = "utf-8"   ' read as UTF-8
        .Open
        .LoadFromFile file_path
        output = .ReadText
        .Close
    End With
    Set stm = Nothing
      
    get_txt_from_file = output
    
outro:
    utilities.call_stack_remove_last_item
    Exit Function
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Function
Public Function json_to_col_of_dicts(ByVal str_path As String) As Collection
    Const proc_name As String = "utilities.json_to_col_of_dicts"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim col_output As Collection
    Dim data_pair
    Dim dict As Scripting.Dictionary
    Dim i As Long
    Dim last_item_no As Long
    Dim str_json As String
    Dim dict_key As String
    Dim dict_value As Variant
    Dim pair_count As Long
    Dim y As Long
    Dim workspace_list
    Dim workspace_data
    
    str_json = utilities.get_txt_from_file(str_path)
    str_json = Left(str_json, Len(str_json) - 1)
    str_json = Right(str_json, Len(str_json) - 1)
    str_json = Replace(str_json, vbNewLine, "")
    str_json = Replace(str_json, vbLf, "")
    str_json = Replace(str_json, vbTab, "")
    
    Do While InStr(str_json, "  ") > 0
        str_json = Replace(str_json, "  ", " ")
    Loop
    
    workspace_list = Split(str_json, "}, {")
    
    workspace_list(0) = Right(workspace_list(0), Len(workspace_list(0)) - 2)
    last_item_no = UBound(workspace_list)
    workspace_list(last_item_no) = Left(workspace_list(last_item_no), Len(workspace_list(last_item_no)) - 2)
    
    Set col_output = New Collection
    For i = 0 To last_item_no
        workspace_data = Split(workspace_list(i), ",")
        
        Set dict = New Scripting.Dictionary
        For y = 0 To UBound(workspace_data) - 1
            data_pair = Split(workspace_data(y), ":")
            
            dict_key = Replace(data_pair(0), """", "")
            dict_key = Trim(dict_key)
            
            dict_value = Replace(data_pair(1), """", "")
            dict_value = Trim(dict_value)
            
            If dict_value = "true" Then
                dict_value = True
            ElseIf dict_value = "false" Then
                dict_value = False
            End If
            
            dict(dict_key) = dict_value
        Next y
        col_output.Add dict
    Next i
    
    Set json_to_col_of_dicts = col_output
    
outro:
    utilities.call_stack_remove_last_item
    Exit Function
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Function
Public Sub log_change(ByRef log_object As utilities.typ_log_object)
    Const proc_name As String = "utilities.log_change"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim log_cmd As ADODB.Command
    
    Set log_cmd = utilities.log_change__create_command
    log_cmd.ActiveConnection = Load.conn
    
    With log_cmd
        .Parameters("p_app_continent").Value = Load.system_info.app_continent
        .Parameters("p_app_name").Value = Load.system_info.app_name
        .Parameters("p_change_source").Value = log_object.change_source
        
        'changer_id
        If log_object.changer_id = "" Then
            .Parameters("p_changer_id").Value = Environ("username")
        Else
            .Parameters("p_changer_id").Value = log_object.changer_id
        End If
        
        .Parameters("p_comment").Value = log_object.comment
        .Parameters("p_data_set_id").Value = log_object.data_set
        
        'deal_id
        .Parameters("p_deal_id").Value = Null
        If log_object.deal_id <> 0 Then .Parameters("p_deal_id").Value = log_object.deal_id
        
        .Parameters("p_event_id").Value = log_object.event_id
        .Parameters("p_executed_sql").Value = log_object.executed_sql
        .Parameters("p_field_name").Value = log_object.field_name
        .Parameters("p_new_value").Value = CStr(log_object.new_value)
        .Parameters("p_operation_type").Value = log_object.operation_type
        
        'policy_id
        .Parameters("p_policy_id").Value = Null
        If log_object.policy_id <> 0 Then .Parameters("p_policy_id").Value = log_object.policy_id
        
        .Parameters("p_record_id").Value = log_object.record_id
        
        'security_id
        .Parameters("p_security_id").Value = Null
        If log_object.security_id <> 0 Then .Parameters("p_security_id").Value = log_object.security_id
        
        .Execute , , adExecuteNoRecords
        
    End With
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "deal_id = " & Nz(log_object.deal_id, 0) & ", event_id = " & log_object.event_id, "", True
    Resume outro
End Sub
Public Sub log_change__field_change( _
    ByRef log_obj_base As utilities.typ_log_object, _
    ByVal field_name As String, _
    ByVal new_val As Variant, _
    ByVal data_set As String _
)
    Const proc_name As String = "utilities.log_change__field_change"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0

    Dim log_obj As utilities.typ_log_object

    log_obj = log_obj_base   ' copy base metadata (event_id, deal_id, etc.)

    log_obj.field_name = field_name
    log_obj.new_value = new_val
    log_obj.data_set = data_set
    
    log_change log_obj
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Sub
Public Function log_change__create_command() As ADODB.Command
    Const proc_name As String = "utilities.log_change__create_command"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    cmd.CommandText = _
    "INSERT INTO log_data_t (" _
        & "app_continent" _
        & ", app_name" _
        & ", change_source" _
        & ", changer_id" _
        & ", comment" _
        & ", data_set_id" _
        & ", deal_id" _
        & ", event_id" _
        & ", executed_sql" _
        & ", field_name" _
        & ", new_value" _
        & ", operation_type" _
        & ", policy_id" _
        & ", record_id" _
        & ", security_id" _
    & ") VALUES (" _
        & " ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?" _
    & ")"
    
    'Parameters must be added in the same order as the ? placeholders
    With cmd.Parameters
        .Append cmd.CreateParameter("p_app_continent", adVarWChar, adParamInput, 50)
        .Append cmd.CreateParameter("p_app_name", adVarWChar, adParamInput, 50)
        .Append cmd.CreateParameter("p_change_source", adVarWChar, adParamInput, 255)
        .Append cmd.CreateParameter("p_changer_id", adVarWChar, adParamInput, 100)
        .Append cmd.CreateParameter("p_comment", adLongVarWChar, adParamInput, -1)
        .Append cmd.CreateParameter("p_data_set_id", adVarWChar, adParamInput, 100)
        .Append cmd.CreateParameter("p_deal_id", adInteger, adParamInput)
        .Append cmd.CreateParameter("p_event_id", adVarWChar, adParamInput, 36)
        .Append cmd.CreateParameter("p_executed_sql", adLongVarWChar, adParamInput, -1)
        .Append cmd.CreateParameter("p_field_name", adVarWChar, adParamInput, 100)
        .Append cmd.CreateParameter("p_new_value", adVarWChar, adParamInput, 255)
        .Append cmd.CreateParameter("p_operation_type", adVarWChar, adParamInput, 20)
        .Append cmd.CreateParameter("p_policy_id", adInteger, adParamInput)
        .Append cmd.CreateParameter("p_record_id", adInteger, adParamInput)
        .Append cmd.CreateParameter("p_security_id", adInteger, adParamInput)
    End With
    
    Set log_change__create_command = cmd

outro:
    utilities.call_stack_remove_last_item
    Exit Function
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "", "", "", True
    Resume outro
End Function

Function get_event_id() As String
    Dim g As GUID
    Dim ret As LongPtr
    
    ret = CoCreateGuid(g)
    
    If ret = 0 Then
        get_event_id = _
            LCase$( _
            Right$("00000000" & Hex$(g.Data1), 8) & "-" & _
            Right$("0000" & Hex$(g.Data2), 4) & "-" & _
            Right$("0000" & Hex$(g.Data3), 4) & "-" & _
            Right$("00" & Hex$(g.Data4(0)), 2) & _
            Right$("00" & Hex$(g.Data4(1)), 2) & "-" & _
            Right$("00" & Hex$(g.Data4(2)), 2) & _
            Right$("00" & Hex$(g.Data4(3)), 2) & _
            Right$("00" & Hex$(g.Data4(4)), 2) & _
            Right$("00" & Hex$(g.Data4(5)), 2) & _
            Right$("00" & Hex$(g.Data4(6)), 2) & _
            Right$("00" & Hex$(g.Data4(7)), 2))
    Else
        get_event_id = ""
    End If
End Function
Function get_html(ByVal input_url As String) As String
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", input_url, False
        .Send
        get_html = .responseText
    End With
End Function
Public Sub call_stack_add_item(ByVal input_proc_name As String)
    If Load.call_stack = "" Then
        Load.event_id = utilities.get_event_id
    End If
    
    Load.call_stack = Load.call_stack & vbNewLine & Time & " " & input_proc_name
End Sub
Public Sub call_stack_remove_last_item(Optional ByVal end_of_event As Boolean)
    Dim pos As Long
    
    If end_of_event = True Then
        Load.event_id = ""
        Load.call_stack = ""
        open_forms.working_on_it_f__close
    Else
        pos = InStrRev(Load.call_stack, vbNewLine)
        If pos > 0 Then Load.call_stack = Left(Load.call_stack, pos - 1)
    End If
End Sub
Public Function expand_command(ByVal cmd As ADODB.Command) As String
    Dim i As Long
    Dim p As ADODB.Parameter
    Dim str_sql As String
    Dim val As String

    str_sql = cmd.CommandText
    For i = 0 To cmd.Parameters.Count - 1
        val = cmd.Parameters(i).Value
        If VarType(val) = vbString Then
            val = "'" & Replace(val, "'", "''") & "'"
        End If
        str_sql = Replace(str_sql, "?", val, , 1) ' replace first "?" only
    Next i
    
    expand_command = str_sql
End Function
Public Sub update_target_sub_sector_list(ByVal target_super_sector_id As Long _
, ByVal target_sub_sector_field_name As String _
, ByVal calling_form As String)
    Const proc_name As String = "utilities.update_target_sub_sector_list"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    'populate target_sub_sector_id based on update of target_super_sector_id
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    Dim target_super_setor_id As Long
    
    If CurrentProject.AllForms(calling_form).IsLoaded = False Then
        GoTo outro
    End If
    
    'remove existing items and add new ones
    With Forms(calling_form).Controls(target_sub_sector_field_name)
        Do While .ListCount > 0
            .RemoveItem (0)
        Loop
        
        str_sql = "SELECT sector_id id, sector_name menu_item FROM " & sources.sectors_table _
        & " WHERE parent_sector_id = " & target_super_sector_id & " ORDER BY sector_name"
        
        Set rs = utilities.create_adodb_rs(conn, str_sql)
        'rs. open
            Do While rs.EOF = False
                .AddItem rs!id & ";'" & rs!menu_item & "'"
                rs.MoveNext
            Loop
        rs.Close
    End With

outro:
    utilities.call_stack_remove_last_item False
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "" _
    , "target_sub_sector_field_name = " & target_sub_sector_field_name & vbNewLine & "target_super_sector_id = " & target_super_sector_id _
    , "", True
    
    Resume outro
End Sub
Public Function create_adodb_rs(ByVal conn As ADODB.Connection, ByVal str_sql As String) As ADODB.Recordset
    Const proc_name As String = "create_adodb_rs"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
        With rs
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockReadOnly
            Set .ActiveConnection = conn
            .Source = str_sql
            .Open
            Set .ActiveConnection = Nothing
        End With
    Set create_adodb_rs = rs
    Set rs = Nothing
        
outro:
    utilities.call_stack_remove_last_item False
    Exit Function
err_handler:
    Debug.Print "utilities.create_adodb_rs: Could not create recordset. str_sql = " & str_sql
    Resume outro
End Function
Public Function generate_sql_date(ByVal input_date As Date) As String
    Dim str_day As String
    Dim str_month As String
    
    If Day(input_date) < 10 Then
        str_day = "0" & CStr(Day(input_date))
    Else
        str_day = CStr(Day(input_date))
    End If
    If Month(input_date) < 10 Then
        str_month = "0" & CStr(Month(input_date))
    Else
        str_month = CStr(Month(input_date))
    End If
    generate_sql_date = CStr(Year(input_date)) & "-" & str_month & "-" & str_day
End Function
Public Function generate_sql_date_2(ByVal input_day As Integer, _
    ByVal input_month As Long, _
    ByVal input_year As Long _
) As String
    Dim str_day As String, str_month As String
    If input_day < 10 Then
        str_day = "0" & CStr(input_day)
    Else
        str_day = CStr(input_day)
    End If
    If input_month < 10 Then
        str_month = "0" & CStr(input_month)
    Else
        str_month = CStr(input_month)
    End If
    generate_sql_date_2 = CStr(input_year) & "-" & str_month & "-" & str_day
End Function

Public Function date_to_weekday(ByVal input_date As Date) As String
    Dim i As Integer, week_day As String
    i = Weekday(input_date)
    If i = 1 Then
        week_day = "Sunday"
    ElseIf i = 2 Then
        week_day = "Monday"
    ElseIf i = 3 Then
        week_day = "Tuesday"
    ElseIf i = 4 Then
        week_day = "Wednesday"
    ElseIf i = 5 Then
        week_day = "Thursday"
    ElseIf i = 6 Then
        week_day = "Friday"
    ElseIf i = 7 Then
        week_day = "Saturday"
    End If
    date_to_weekday = week_day
End Function
Public Function twips_converter(ByVal input_number, ByVal inch_or_cm) As Long
    If inch_or_cm = "inch" Then
        twips_converter = input_number * 1440
    ElseIf inch_or_cm = "cm" Then
        twips_converter = input_number * 1440 / 2.54
    Else
        twips_converter = -1
    End If
End Function
Public Function english_month_name(ByVal month_number As Integer) As String
    Dim output As String
    If month_number = 1 Then output = "January"
    If month_number = 2 Then output = "February"
    If month_number = 3 Then output = "March"
    If month_number = 4 Then output = "April"
    If month_number = 5 Then output = "May"
    If month_number = 6 Then output = "June"
    If month_number = 7 Then output = "July"
    If month_number = 8 Then output = "August"
    If month_number = 9 Then output = "September"
    If month_number = 10 Then output = "October"
    If month_number = 11 Then output = "November"
    If month_number = 12 Then output = "December"
    
    english_month_name = output
End Function
Function convert_long_color_to_rgb(color_value As Long) As String
    Dim Red As Long, Green As Long, Blue As Long
    Red = color_value Mod 256
    Green = ((color_value - Red) / 256) Mod 256
    Blue = ((color_value - Red - (Green * 256)) / 256 / 256) Mod 256
    
    convert_long_color_to_rgb = "RGB(" & _
                    Red & ", " & _
                    Green & ", " & _
                    Blue & ")"
End Function
Public Function encrypt_string(ByVal input_string As String, _
    ByVal salt_factor As Integer, _
    ByVal offset_factor As Integer _
) As String
    Const proc_name As String = "utilities.encrypt_string"
    utilities.call_stack_add_item proc_name
    
    Dim i As Long
    Dim input_collection As New Collection
    Dim output_string As String
    Dim random_char As String
    Dim random_chr As Long
    Dim temp_output As New Collection
    Dim y As Long
    
    For i = 1 To Len(input_string)
        input_collection.Add Mid(input_string, i, 1)
    Next i
    
    For i = 1 To Len(input_string) - 1
        temp_output.Add ChrW(Asc(input_collection(i)) + offset_factor)
        For y = 1 To salt_factor
            random_char = Int(Rnd * (254 - 174))
            temp_output.Add ChrW(random_char)
        Next y
    Next i
    temp_output.Add Chr(Asc(input_collection(input_collection.Count)) + offset_factor)
    
    For i = 1 To temp_output.Count
        output_string = output_string & temp_output(i)
    Next i
    encrypt_string = output_string
outro:
    utilities.call_stack_remove_last_item False
    Exit Function
End Function
Public Function weekday_converter(ByVal input_date As Date) As String
    Dim int_day As Long
    int_day = Weekday(input_date, vbSunday)
    Dim output As String
    If int_day = 1 Then
        output = "Sunday"
    ElseIf int_day = 2 Then
        output = "Monday"
    ElseIf int_day = 3 Then
        output = "Tuesday"
    ElseIf int_day = 4 Then
        output = "Wednesday"
    ElseIf int_day = 5 Then
        output = "Thursday"
    ElseIf int_day = 6 Then
        output = "Friday"
    ElseIf int_day = 7 Then
        output = "Saturday"
    End If
    weekday_converter = output
End Function

Public Function decrypt_string(ByVal input_string As String, ByVal salt_factor As Integer, ByVal offset_factor As Integer) As String
    Dim pw_int As String, i As Integer, pw As String
    pw_int = pw_int & Mid(input_string, 1, 1)
    For i = 2 + salt_factor To Len(input_string) Step salt_factor + 1
        pw_int = pw_int & Mid(input_string, i, 1)
    Next i
    Dim int_letter As Integer
    pw = ""
    For i = 1 To Len(pw_int)
        int_letter = Asc(Mid(pw_int, i, 1))
        pw = pw & ChrW(int_letter - offset_factor)
    Next i
    decrypt_string = pw
outro:
End Function
Public Function remove_illegal_characters(ByVal input_string As String) As String
    Dim str_output As String
    str_output = input_string
    str_output = Replace(str_output, "\", "")
    str_output = Replace(str_output, "/", "")
    str_output = Replace(str_output, ";", "")
    str_output = Replace(str_output, ":", "")
    str_output = Replace(str_output, "%", "")
    str_output = Replace(str_output, "!", "")
    str_output = Replace(str_output, "?", "")
    str_output = Replace(str_output, "'", "")
    
    remove_illegal_characters = str_output
End Function

Public Function get_public_ip()
    Load.call_stack = Load.call_stack & vbNewLine & "utilities.get_public_ip"
    Dim url As String, ip_adr As String
    With CreateObject("MSXML2.XMLHTTP.6.0")
        url = "https://checkip.amazonaws.com/"
        .Open "GET", url, False
        .Send
        ip_adr = .responseText
    Dim reg_exp As Object

    Set reg_exp = CreateObject("vbscript.regexp")
        If .status = 200 Then
            With reg_exp
                .Pattern = "\s"
                .MultiLine = True
                .global = True
                get_public_ip = .Replace(ip_adr, vbNullString)
            End With
        Else
            get_public_ip = "-1"
        End If
    End With
End Function

Public Sub paint_control(ByVal form_name As String, ByVal col_input_controls As Collection)
    Const proc_name As String = "utilities.paint_control_name"
    utilities.call_stack_add_item proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0

    If CurrentProject.AllForms(form_name).IsLoaded = False Then
        GoTo outro
    End If
    
    Dim input_control As cls_field
    For Each input_control In col_input_controls
        With Forms(form_name).Controls(input_control.field_name)
            If input_control.field_bg_color <> -1 Then
                .BackColor = input_control.field_bg_color
            End If
            If input_control.field_caption <> "-1" Then
                .Caption = input_control.field_caption
            End If
            If input_control.font_size <> -1 Then
                .FontSize = input_control.font_size
            End If
            If input_control.font_color <> -1 Then
                .ForeColor = input_control.font_color
            End If
            If input_control.field_height <> -1 Then
                .Height = input_control.field_height
            End If
            If input_control.field_left <> -1 Then
                .Left = input_control.field_left
            End If
            If input_control.field_top <> -1 Then
                .Top = input_control.field_top
            End If
            .Visible = input_control.field_visible
            If input_control.field_width <> -1 Then
                .Width = input_control.field_width
            End If
            If input_control.font_alignment <> -1 Then
                .TextAlign = input_control.font_alignment
            End If
            If input_control.font_type <> "-1" Then
                .FontName = input_control.font_type
            End If
            If input_control.field_type = Load.field_type.text_field Then
                If input_control.margin_internal_bottom <> -1 Then
                    .BottomMargin = input_control.margin_internal_bottom
                End If
                If input_control.margin_internal_left <> -1 Then
                    .LeftMargin = input_control.margin_internal_left
                End If
                If input_control.margin_internal_top <> -1 Then
                    .TopMargin = input_control.margin_internal_top
                End If
                If input_control.field_value <> "-1" Then
                    .Value = input_control.field_value
                End If
                If input_control.field_name_in_recordset <> "-1" Then
                    .ControlSource = input_control.field_name_in_recordset
                End If
            End If
        End With
    Next input_control
    
outro:
    utilities.call_stack_remove_last_item
    Exit Sub
err_handler:
    Central.err_handler proc_name, Err.Number, Err.Description, "input_control.field_name = " & Nz(input_control.field_name, ""), "", "", True
    Resume outro
End Sub

Public Function convert_number_to_figure(ByVal input_number As Long) As String
    If input_number = 1 Then convert_number_to_figure = "a"
    If input_number = 2 Then convert_number_to_figure = "b"
    If input_number = 3 Then convert_number_to_figure = "c"
    If input_number = 4 Then convert_number_to_figure = "d"
    If input_number = 5 Then convert_number_to_figure = "e"
    If input_number = 6 Then convert_number_to_figure = "f"
    If input_number = 7 Then convert_number_to_figure = "g"
    If input_number = 8 Then convert_number_to_figure = "h"
    If input_number = 9 Then convert_number_to_figure = "i"
    If input_number = 10 Then convert_number_to_figure = "j"
    If input_number = 11 Then convert_number_to_figure = "k"
    If input_number = 12 Then convert_number_to_figure = "l"
    If input_number = 13 Then convert_number_to_figure = "m"
    If input_number = 14 Then convert_number_to_figure = "n"
    If input_number = 15 Then convert_number_to_figure = "o"
    If input_number = 17 Then convert_number_to_figure = "p"
    If input_number = 18 Then convert_number_to_figure = "q"
    If input_number = 19 Then convert_number_to_figure = "r"
    If input_number = 20 Then convert_number_to_figure = "s"
    If input_number = 21 Then convert_number_to_figure = "t"
    If input_number = 22 Then convert_number_to_figure = "u"
    If input_number = 23 Then convert_number_to_figure = "v"
    If input_number = 24 Then convert_number_to_figure = "x"
    If input_number = 25 Then convert_number_to_figure = "y"
    If input_number = 26 Then convert_number_to_figure = "z"
    If input_number = 27 Then convert_number_to_figure = "aa"
    If input_number = 28 Then convert_number_to_figure = "ab"
    If input_number = 29 Then convert_number_to_figure = "ac"
    If input_number = 30 Then convert_number_to_figure = "ad"
    If input_number = 31 Then convert_number_to_figure = "ae"
    If input_number = 32 Then convert_number_to_figure = "af"
    If input_number = 33 Then convert_number_to_figure = "ag"
    If input_number = 34 Then convert_number_to_figure = "ah"
    If input_number = 35 Then convert_number_to_figure = "ai"
    If input_number = 36 Then convert_number_to_figure = "aj"
    If input_number = 37 Then convert_number_to_figure = "ak"
    If input_number = 38 Then convert_number_to_figure = "al"
    If input_number = 39 Then convert_number_to_figure = "am"
    If input_number = 40 Then convert_number_to_figure = "an"
    If input_number = 41 Then convert_number_to_figure = "ao"
    If input_number = 42 Then convert_number_to_figure = "ap"
End Function