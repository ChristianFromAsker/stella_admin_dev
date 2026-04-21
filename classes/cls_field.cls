Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public above_field As cls_field
Public border_color As Long

Public calculated_at_runtime As Boolean
Public colour_hover As Long

Public data_table As String

Public excel_column_width As Long
Public excel_horizontal_alignment As Long

Public field_bg_color As Long
Public field_caption As String
Public field_caption_us As String
Public field_height As Long

Public field_left As Long
Public field_name As String
Public field_name_add_many_deals_f As String
Public field_name_deals_v As String
Public field_name_in_recordset As String
Public field_name_in_table As String

Public field_top As Long
Public field_type As String

Public field_value As Variant
Public field_visible As Boolean

Public field_width As Long

Public font_alignment As Long
Public font_color As Long
Public font_color_hover As Long
Public font_color_pressed As Long
Public font_is_bold As Boolean
Public font_size As Long
Public font_type As String

Public has_help_text As Boolean

Public id As Long
Public is_active As Boolean
Public is_collapsed As Boolean
Public is_mandatory As Boolean

Public margin_external_bottom As Long
Public margin_external_left As Long
Public margin_external_right As Long
Public margin_external_top As Long
Public margin_internal_bottom As Long
Public margin_internal_left As Long
Public margin_internal_right As Long
Public margin_internal_top As Long

Public section_height As Long
Public top_margin As Long

Public user_has_expanded_section As Boolean

Public visible_in_america As Boolean
Public visible_in_europe As Boolean

'Access cannot set left, right, top and bottom border. Hence, I (CK) have created an option to simulate only having a bottom border.
Public has_line As Boolean
Public line_name As String

Private Sub Class_Initialize()
    With Me
        .colour_hover = -1
        .field_bg_color = -1
        .field_caption = "-1"
        .field_caption_us = "-1"
        .field_height = -1
        .field_left = -1
        .field_name_in_recordset = "-1"
        .field_name_in_table = "-1"
        .field_top = -1
        .field_type = "-1"
        .field_value = "-1"
        .field_width = -1
        .font_alignment = -1
        .font_color = -1
        .font_size = -1
        .font_type = "-1"
        .margin_external_bottom = -1
        .margin_external_left = 0
        .margin_external_right = -1
        .margin_external_top = -1
        .margin_internal_bottom = -1
        .margin_internal_left = -1
        .margin_internal_right = -1
        .margin_internal_top = -1
        .visible_in_america = True
        .visible_in_europe = True
    End With
End Sub

Public Function field_bottom()
    field_bottom = Me.field_top + Me.field_height
End Function
Public Function field_right()
    Dim margin_external_left As Long
    margin_external_left = 0
    field_right = Me.field_left + Me.field_width
End Function
Public Function copy_field() As cls_field
    Dim output_field As cls_field
    With output_field
        .above_field = Me.above_field
        .border_color = Me.border_color
        .data_table = Me.data_table
        .field_bg_color = Me.field_bg_color
        .field_caption = Me.field_caption
        .field_caption_us = Me.field_caption_us
        .field_height = Me.field_height
        .field_left = Me.field_left
        .field_name = Me.field_name
        .field_name_in_table = Me.field_name_in_table
        .field_top = Me.field_top
        .field_type = Me.field_type
        .field_value = Me.field_value
        .field_visible = Me.field_visible
        .field_width = Me.field_width
        .visible_in_america = Me.visible_in_america
        .visible_in_europe = Me.visible_in_europe
    End With
    Set copy_field = output_field
End Function