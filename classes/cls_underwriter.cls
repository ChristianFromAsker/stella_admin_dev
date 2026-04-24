Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Public admin_access As Integer
Public budget_home As Variant
Public can_change_general_id As Integer
Public can_change_jurisdictions_id As Integer
Public can_change_uws_id As Integer
Public can_change_budget_home_id As Integer
Public full_name As String
Public is_dev_id As Integer
Public is_regional_lead_id As Integer
Public initials As String
Public nick As String
Public user_name As String
Public uw_id As Long
Public windows_user_name As String
Private Sub Class_Initialize()
    windows_user_name = Environ("Username")
End Sub
Public Sub init_uw(Optional input_uw_id As Long)
    Dim str_sql As String, rs As ADODB.Recordset, str_where As String
    
    str_where = "user_name = '" & windows_user_name & "'"
    If windows_user_name = "chris" Then str_where = "user_name = 'christian.kartnes'"
    
    If input_uw_id <> 0 Then
        str_where = "uw_id  = " & input_uw_id
    End If
    
    str_sql = "SELECT uw_id, uw_initials, uw_name, user_name, has_admin_access_id, nickname, budget_home_id " _
        & ", can_change_general_id, can_change_jurisdictions_id, can_change_uws_id, can_change_budget_home_id, is_dev_id, is_regional_lead_id" _
        & " FROM " & sources.underwriters_table _
        & " WHERE " & str_where
        
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        admin_access = rs!has_admin_access_id
        budget_home = Nz(rs!budget_home_id, "-1")
        user_name = rs!user_name
        full_name = rs!uw_name
        uw_id = rs!uw_id
        initials = rs!uw_initials
        nick = Nz(rs!nickname, full_name)
        can_change_general_id = rs!can_change_general_id
        can_change_jurisdictions_id = rs!can_change_jurisdictions_id
        can_change_uws_id = rs!can_change_uws_id
        can_change_budget_home_id = rs!can_change_budget_home_id
        is_dev_id = rs!is_dev_id
        is_regional_lead_id = rs!is_regional_lead_id
    rs.Close: Set rs = Nothing
End Sub