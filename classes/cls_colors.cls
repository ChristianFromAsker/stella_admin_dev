Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public off_white As Variant
Public light_red As Variant
Public yellow As Variant
Public light_green As Variant
Public white As Variant
Public very_light_blue As Variant
Public very_dark_blue As Variant
Public very_light_green As Variant
Public black As Variant
Public dark_blue As Variant
Public active_filter_bg As Variant
Public inactive_filter_bg As Variant
Public inactive_navins_home As Long
Public active_navins_home As Long

'by function
Public command_button As Long
Public Sub init()
    With Me
        .off_white = 15529969
        .light_red = 11389944
        yellow = RGB(255, 204, 0)
        light_green = RGB(197, 224, 180)
        white = RGB(255, 255, 255)
        very_light_blue = RGB(189, 215, 238)
        very_dark_blue = RGB(34, 42, 53)
        black = RGB(0, 0, 0)
        dark_blue = RGB(170, 200, 210)
        very_light_green = RGB(241, 247, 236)
        active_filter_bg = RGB(0, 168, 45)
        inactive_filter_bg = white
        active_navins_home = RGB(1, 121, 111)
        inactive_navins_home = RGB(255, 255, 255)
        
        'by function
        .command_button = RGB(1, 121, 111)
    End With
End Sub
Private Sub Class_Initialize()
    Me.init
End Sub