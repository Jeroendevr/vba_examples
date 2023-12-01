Public Function getreq(url As String)
    Dim req As WinHttpRequest
    Dim JsonString As String
    Dim jp As Object
    Dim resp As String
    Dim errorstring As String

    Set req = New WinHttpRequest
    req.Open "GET", url
    req.setRequestHeader "Content-Type", "multipart/form-data"
    req.setRequestHeader "Accept", "application/xml"
    On Error GoTo errhand:
        req.Send
        resp = req.ResponseText
        If resp = "Internal Server Error" Then
            resp = "{'error': 'Internal server error'}"
        End If
        getreq = resp
    Exit Function

errhand:

    Select Case Err.Number
        Case -2147012894 'Code for Timeout
            getreq = "{'error': 'Request timeout'}"
        Case -2147012891 'Code for Invalid URL
            getreq = "{'error': 'Bad url'}"
        Case -2147012867 'Code for Invalid URL
            getreq = "{'error': 'Cannot establish connection'}"
        Case Else 'Add more Errorcodes here if wanted
            errorstring = "Errornumber: " & Err.Number & vbNewLine & "Errordescription: " & Error(Err.Number)
            getreq = "{'error': '" & errorstring & "'}"
    End Select
End Function

Sub apicall()
    Dim url As String
    url = "https://{url}"
    
    Dim response As Object
    Set response = JsonConverter.ParseJson(getreq_av(url))
    
    Dim prices As Dictionary
    Set prices = response("{key}")
    
End Sub
