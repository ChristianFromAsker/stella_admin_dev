Option Compare Database
Option Explicit
#If VBA7 Then
    #If Win64 Then
        Public Type POINTAPI
            x As Long
            y As Long
        End Type
        Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hWnd As LongPtr, lpPoint As POINTAPI) As Long
        Declare PtrSafe Function apiGetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
        Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" ( _
            ByVal hWnd As LongPtr, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal nWidth As Long, _
            ByVal nHeight As Long, _
            ByVal bRepaint As Long _
         ) As Long
         Declare PtrSafe Function apiGetParent Lib "user32" Alias "GetParent" (ByVal hWnd As LongPtr) As Long
         Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
    #Else
        Declare Function apiGetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
        Declare Function apiGetParent Lib "user32" Alias "GetParent" (ByVal hWnd As LongPtr) As Long
        Declare Function apiMoveWindow Lib "user32" Alias "MoveWindow" ( _
            ByVal hWnd As LongPtr, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal nWidth As Long, _
            ByVal nHeight As Long, _
            ByVal bRepaint As Long _
         ) As Long
    #End If
#End If
    ' This is used to open the intranet from the "Intranet" button on the main menu

    'Private Declare Function ShellExecute _
        '  Lib "shell32.dll" Alias "ShellExecuteA" ( _
        '  ByVal hWnd As Long, _
        '  ByVal Operation As String, _
        '  ByVal Filename As String, _
        '  Optional ByVal Parameters As String, _
        '  Optional ByVal Directory As String, _
        '  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
    '  ) As Long

Function AccessMoveSize(iX As Integer, iY As Integer, iWidth As Integer, iHeight As Integer)
    apiMoveWindow GetAccesshWnd(), iX, iY, iWidth, iHeight, True
End Function
Function GetAccesshWnd()
    Dim hWnd As LongPtr
    Dim hWndAccess As LongPtr
    hWnd = apiGetActiveWindow()
    hWndAccess = hWnd
    While hWnd <> 0
        hWndAccess = hWnd
        hWnd = apiGetParent(hWnd)
    Wend
    GetAccesshWnd = hWndAccess
End Function