Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public button_name As String
Public button_value As Boolean

Public inactive_bg_color As Long
Public active_bg_color As Long
Public default_bg_color As Long
Public current_bg_color As Long

Public active_caption As String
Public inactive_caption As String
Public default_caption As String
Public current_caption As String

Public active_font_color As Long
Public inactive_font_color As Long
Public default_font_color As Long
Public current_font_color As Long

Public menu_id As Long

Public Sub Class_Initialize()
    active_bg_color = Load.colors.active_navins_home
    inactive_bg_color = Load.colors.inactive_navins_home
    active_font_color = Load.colors.white
    inactive_font_color = Load.colors.black
    button_value = True
    current_bg_color = Load.colors.active_navins_home
    current_font_color = Load.colors.white
End Sub