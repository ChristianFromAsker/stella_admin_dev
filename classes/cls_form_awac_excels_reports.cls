Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Public payment_report As String
Public written_report As String
Public payment_converted_file_path As String
Public payment_original_file_path As String
Public payment_original_file_name As String
Public v5_original_incl_top10_rows_file_path As String
Public v5_converted_file_path As String
Public v5_original_file_path As String
Public v5_renamed_file_path As String
Public v5_original_file_name As String

Private Sub Class_Initialize()
    payment_report = "payment_report"
    written_report = "written_report"
End Sub