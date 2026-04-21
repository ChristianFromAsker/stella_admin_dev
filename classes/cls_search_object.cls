Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
' Public status(1 To 10) As Integer
Public date_type As Integer
Public Sector As Integer
Public Target As String
' Public broker(1 To 5) As Integer
' Public spa_law(1 To 3) As Integer
Public deal_type As Integer
Private status As Variant
Private Sub Class_Initialize()
    Dim arr_status(1 To 10) As Boolean
    For i = 1 To 6
        arr_status(i) = True
    Next i
    For i = 7 To 10
        arr_status(i) = False
    Next i
    status = arr_status
End Sub
Property Let give_status(input_status As Variant)
    status = input_status
End Property
Property Get get_status(i As Integer) As Variant
    get_status = status(i)
End Property