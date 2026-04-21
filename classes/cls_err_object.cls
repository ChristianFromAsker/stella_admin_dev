Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'test
Public call_stack As String
Public system_error_text As String
Public system_error_code As String
Public stella_error_text As String
Public routine_name As String
Public params As String
Public milestone As String
Public show_error_msg As Boolean

Private Sub placeholder()

'code to be copied at start of subs
    'load.call_stack = Me.Name & "." & Screen.ActiveControl.Name
    Load.call_stack = Load.call_stack & vbNewLine & "[proc name]"
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0

'code to be copied above err_handler
outro:
    'load.call_stack = "" ' if initiating sub
    'open_forms.working_on_it_f__close ' if initiating sub
    Exit Sub

'code to be inserted in err_handler of subs
err_handler:
    Dim err_object As cls_err_object
    Set err_object = New cls_err_object
    With err_object
        .routine_name = ""
        .milestone = ""
        .params = ""
        .system_error_code = Err.Number
        .system_error_text = Err.Description
        .show_error_msg = True
        .send_error err_object
    End With
    GoTo outro
    
End Sub

Public Sub send_error(ByVal err_object As cls_err_object)
    Load.call_stack = Load.call_stack & vbNewLine & "cls_err_object.send_error"
    Dim str_sql As String, uw_name As String, app_name As String, app_continent As String
    On Error Resume Next
    If Load.is_debugging = True Then On Error GoTo 0
    
    If Load.is_debugging = True Then
        GoTo outro
    End If
        
    app_name = "-1"
    app_name = Load.system_info.app_name
    app_continent = "-1"
    app_continent = Load.system_info.app_continent
    
    With err_object
        str_sql = "INSERT INTO log_errors_t (system_error_text" _
            & ", system_error_code" _
            & ", stella_error_text" _
            & ", routine_name" _
            & ", call_stack" _
            & ", params" _
            & ", milestone" _
            & ", uw_name" _
            & ", app_name" _
            & ", file_path" _
            & ", app_continent)" _
            & " VALUES (" _
            & "'" & .system_error_text & "'" _
            & ", '" & .system_error_code & "'" _
            & ", '" & .stella_error_text & "'" _
            & ", '" & .routine_name & "'" _
            & ", '" & Load.call_stack & "'" _
            & ", '" & .params & "'" _
            & ", '" & .milestone & "'" _
            & ", '" & Environ("Username") & "'" _
            & ", '" & app_name & "'" _
            & ", '" & Replace(CurrentProject.FullName, "\", "\\") & "'" _
            & ", '" & app_continent & "'" _
            & ")"
        End With
    
    conn.Execute str_sql
    
    Err.Clear
    
    If Me.show_error_msg = True Then
        MsgBox Load.system_info.error_instruction & vbNewLine & vbNewLine _
            & "Error number: " & Me.system_error_code & vbNewLine _
            & "Error description: " & Me.system_error_text & vbNewLine _
            & "Where: " & Me.routine_name & vbNewLine _
            & "Parameters: " & Me.params & vbNewLine _
            & "App: " & Load.system_info.app_name, , Load.system_info.error_msg_heading
    End If
outro:
    Exit Sub
    
End Sub