Option Compare Database
Function create_folder_path(ByVal obj_deal As Object)
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    Dim str_sql As String, rs As ADODB.Recordset, folder_name As String, str_year As String
    str_sql = "SELECT * FROM " & Load.sources.menu_list_table & " WHERE menu_id = " & obj_deal.status
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        str_status_path = rs!setting1
    rs.Close
    Set rs = Nothing
    
    If obj_deal.status <= 5 Or obj_deal.status = 481 Then
        str_year = ""
    ElseIf obj_deal.status = 6 Or obj_deal.status = 436 Then
        str_year = obj_deal.inception_year & "\"
    Else
        str_year = obj_deal.create_year & "\"
    End If
    
    folder_name = Paths.find_deal_folder_name(obj_deal)
    If folder_name = "-1" Then
        create_folder_path = "-1"
    Else
        create_folder_path = Load.system_info.system_paths.base_path & str_status_path & str_year & folder_name
    End If
    
outro:
    Exit Function
    
err_handler:
    create_folder_path = "-1"
    
End Function
Public Function IDNumber(ByVal deal_id As Long) As String
    ' This function creates the six digit ID number to be placed at the end of folder names on working.
    Dim str_deal_id As String
    str_deal_id = deal_id

    Do Until Len(str_deal_id) = 6
        str_deal_id = "0" + str_deal_id
    Loop
    
    IDNumber = str_deal_id
End Function

Public Function find_deal_folder_name(ByVal obj_deal As cls_deal) As String
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    ' Finds country path and folder name of deal
    Dim rs As ADODB.Recordset
    Dim str_sql As String
    'Find broker part of deal folder
    Dim str_broker As String
    str_sql = "SELECT * FROM broker_firms_v WHERE id = " & obj_deal.broker_firm
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        str_broker = "(" & rs!short_name & ")"
    rs.Close
    ' Find country abbrevation part of deal folder
    Dim int_menu_item As Integer
    str_sql = "SELECT * FROM " & Load.sources.jurisdictions_table & " WHERE jurisdiction = '" & obj_deal.spa_law & "'"
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        str_country = "(" & rs!abbrevation & ")"
        int_menu_item = rs!working_folder_id
    rs.Close
    str_sql = "SELECT * FROM " & Load.sources.menu_list_table & " WHERE menu_id = " & int_menu_item
    Dim str_country_path As String
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        str_country_path = rs!menu_item & "\"
    rs.Close
    Set rs = Nothing
    find_deal_folder_name = str_country_path & obj_deal.deal_name & " " & str_country & str_broker & " " & Paths.IDNumber(obj_deal.deal_id)

outro:
    Exit Function
    
err_handler:
    find_deal_folder_name = "-1"
    
End Function