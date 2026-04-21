Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
'deal info
Public deal_name As String
Public broker_firm As String
Public broker_firm_hr As String
Public closing_date As Variant
Public spa_law As String
Public Target As String
Public target_naic_code As String
Public target_sic_code As String
Public type_hr As String
' Insurance info
Public broker_commission As Double
Public inception_date As Variant
Public internal_approver_hr As String
Public Layer As Integer
Public layer_hr As String
Public lead_insurer_hr As String
Public policy_no As String
Public policy_period As Integer
Public primary_uw_hr As String
Public program_limit As Long
Public status As Integer
Public status_hr As String
'Figures
Public deal_currency As String
Public EV As String
Public retention As Long
Public drop_end As Long
Public drop_period As Integer
Public Limit As Long
Public premium As Long
Public signing_invoice_amount As Long
'insured info
Public buyer_business_name As String
Public law_firm_1_hr As String
Public insured_name As String
Public insured_zip_city As String
Public insured_street As String
Public insured_region As String
Public insured_country As String
'seller
Public seller_legal_name As String
Public SellerDomicile As String
Public seller_comment As String
' Misc
Public Comments As String
Public create_date As Variant
Public deal_id As String
Public message_board As Variant
Property Get target_path() As String
    target_path = Left(Target, 15)
End Property
Property Get create_year()
    create_year = Year(create_date)
End Property
Property Get inception_year()
    If IsNull(inception_date) = Null Then
        inception_year = 666
    Else
        inception_year = Year(inception_date)
    End If
End Property
Property Get surplus_broker_info()
    ' Only relevant for Stella US
    surplus_broker_info = 666
End Property
Property Get program_summary()
    Dim str_sql As String
    Dim rs As ADODB.Recordset
    str_sql = "SELECT deal_id, program_summary from deals_v WHERE deal_id = " & deal_id
    Set rs = utilities.create_adodb_rs(conn, str_sql)
        program_summary = rs!program_summary
    rs.Close
    Set rs = Nothing
End Property