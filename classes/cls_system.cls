Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public app_name As String
Public ip_current As String
Public error_msg_heading As String
Public error_instruction As String
Public app_continent As String
Public database_name As String
Public is_init As Boolean
Public continents As New cls_system_continents

Public system_paths As New cls_system_paths

Private Sub Class_Initialize()
    Me.init_system
End Sub

Public Sub init_system()
    Dim proc_name As String, str_milestone As String
    proc_name = "cls_system.init_system"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    str_milestone = "1"
    Me.continents.init
    With Me
        .ip_current = utilities.get_public_ip
        
        'try to set app_continent based on ip. Use available drives if not.
        str_milestone = "2"
        If Load.system_info.app_continent = "" Then
            If .ip_current = Me.continents.eurasia_ip Then Load.system_info.app_continent = Me.continents.eurasia
            If .ip_current = Me.continents.america_ip Then .app_continent = Me.continents.americas
        End If
        
        str_milestone = "3"
        If .app_continent = "" Then
            'Most Americans do not have access to F:\
            If Dir("F:\", vbDirectory) = "" Then
                .app_continent = Me.continents.americas
            Else
                .app_continent = Me.continents.eurasia
            End If
        End If
        
        str_milestone = "4"
        .app_name = "Stella Admin"
        .error_msg_heading = "Darn it!"
        .error_instruction = "Something went wrong. Snip this message to Tom (tom.evans@rpgroup.com) or Christian (christian.kartnes@rpgroup.com)" _
            & ", and include what you did just before the error happened."
        
        str_milestone = "5"
        If .app_continent = Me.continents.eurasia Then
            .database_name = "stella_eur"
        ElseIf .app_continent = Me.continents.americas Then
            .database_name = "stella_us"
        End If
        
        str_milestone = "6"
        Me.system_paths.init
        
        .is_init = True
    End With
    
outro:
    Exit Sub
    
err_handler:
    MsgBox "Something went wrong. Try once more. If it fails again, snip this to Christian." & vbNewLine & vbNewLine _
        & "Error number: " & Err.Number & vbNewLine _
        & "Error description: " & Err.Description & vbNewLine _
        & "Where: " & proc_name & vbNewLine _
        & "Milestone: " & str_milestone & vbNewLine _
        & "Parameters: n/a" & vbNewLine _
        & "Call stack: " & Load.call_stack & vbNewLine _
        & "App: Stella Admin", , " Whoopsie daisies (like Hugh Grant in Notting Hill)"
    GoTo outro
End Sub