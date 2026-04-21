Option Compare Database
Option Explicit

Public Sub load_module_speed_test()
    Const test_count As Long = 10
    
    Dim dct_resource As Scripting.Dictionary
    Dim resource_type As AcObjectType
    Dim col_resources As Collection
    Dim i As Long
    Dim name_of_module_to_be_imported As String
    Dim module_file_path As String
    
        'at start of test sequence
    Dim timer_start As Single
    timer_start = Timer
    
    Set col_resources = New Collection
    
    Set dct_resource = New Scripting.Dictionary
    dct_resource.Add "resource_name", "utilities"
    dct_resource.Add "resource_path", "D:\riskpoint\stella\stella_uw_dev\modules\" & dct_resource("resource_name")
    dct_resource.Add "resource_type", acModule
    dct_resource.Add "resource_ending", "bas"
    col_resources.Add dct_resource
        
    For i = 1 To test_count
        For Each dct_resource In col_resources
            Application.LoadFromText dct_resource("resource_type"), dct_resource("resource_name"), dct_resource("resource_path") & "." & dct_resource("resource_ending")
        Next dct_resource
    Next i
    
    'at end of test sequence
    Debug.Print "Test took: " & Timer - timer_start
End Sub
Sub RetrieveAttachments()
    
    Dim olApp As Outlook.Application
    Dim olNs As Outlook.Namespace
    Dim olFolder As Outlook.MAPIFolder
    Dim olMail As Outlook.MailItem
    Dim olAttachment As Outlook.Attachment
    Dim strFolderPath As String
    
    ' Initialize Outlook Application
    Set olApp = New Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    
    ' Access the client's mailbox folder
    Set olFolder = olNs.Folders("bbi").Folders("Inbox")
    
    ' Loop through each email in the folder
    For Each olMail In olFolder.Items
        ' Check if the email is from a specific sender
        
        
            
            Debug.Print olMail.SenderEmailAddress
        
    Next olMail
    
    ' Cleanup
    Set olAttachment = Nothing
    Set olMail = Nothing
    Set olFolder = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
End Sub
Public Sub replace_file_in_use()
    Load.check_conn_and_variables
    Dim fso As Object, source_path As String, target_path As String
    source_path = Load.system_info.system_paths.stella_path & "stable_builds\stella_admin.accdb"
    target_path = Load.system_info.system_paths.stella_path & "published\stella_admin.accdb"
    Set fso = CreateObject("scripting.filesystemobject")
        fso.CopyFile source_path, target_path, True
    Set fso = Nothing
    
End Sub

Public Sub scraping()
    Dim winHttpReq As Object

    Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    Dim url As String
    url = "https://www.google.com/search?q=renewable"
   
    With winHttpReq

        .Open "GET", url, False
        .Send

        'Debug.Print "Response Headers:"
        'Debug.Print .getAllResponseHeaders

        Debug.Print "Response Text:"
        Debug.Print .responseText

    End With

    Set winHttpReq = Nothing
   
End Sub
Sub Gethits()
    Dim url As String
    Dim XMLHTTP As Object, html As Object, objResultDiv As Object, objH3 As Object, link As Object
    Dim var As String
    Dim var1 As Object

    Dim cookie As String
    Dim result_cookie As String
    url = "https://www.google.com/search?q=renewable"
    
    url = "https://www.vg.no"
    
    Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
    XMLHTTP.Open "GET", url, False
    XMLHTTP.setRequestHeader "Content-Type", "text/xml"
    XMLHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; rv:25.0) Gecko/20100101 Firefox/25.0"
    XMLHTTP.Send

    Set html = CreateObject("htmlfile")
    html.Body.innerHTML = XMLHTTP.responseText
    Set objResultDiv = html.getElementById("rso")
    Set var1 = html.getElementById("resultStats")
    
    
    Debug.Print XMLHTTP.responseText

End Sub
Public Sub LinesOfCode()
    Dim vbeModule As Object
    Dim LinesOfCode_local As Long
    LinesOfCode_local = 0
    For Each vbeModule In Application.VBE.ActiveVBProject.VBComponents
        LinesOfCode_local = LinesOfCode_local + vbeModule.CodeModule.CountOfLines
    Next vbeModule
    MsgBox LinesOfCode_local
End Sub