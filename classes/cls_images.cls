Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public MenuMainF As String

Private Sub Class_Initialize()
    Dim proc_name As String, str_milestone As String
    proc_name = "cls_images.Class_Initialize"
    Load.call_stack = Load.call_stack & vbNewLine & proc_name
    On Error GoTo err_handler
    If Load.is_debugging = True Then On Error GoTo 0
    
    MenuMainF = Load.system_info.system_paths.images_path & "main_menu_02_mountains.jpg"
    If Month(Date) > 2 Then MenuMainF = Load.system_info.system_paths.images_path & "main_menu_01_forest.jpg"
    
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