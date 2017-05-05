```vb
Function Download_File(ByVal vWebFile As String, ByVal vLocalFile As String) As Boolean
    Dim oXMLHTTP As Object, i As Long, vFF As Long, oResp() As Byte

    'You can also set a ref. to Microsoft XML, and Dim oXMLHTTP as MSXML2.XMLHTTP
    Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    oXMLHTTP.Open "GET", vWebFile, False         'Open socket to get the website
    oXMLHTTP.Send                                'send request

    'Wait for request to finish
    Do While oXMLHTTP.readyState <> 4
        DoEvents
    Loop

    oResp = oXMLHTTP.responseBody                'Returns the results as a byte array

    'Create local file and save results to it
    Dim oStream As Object
    If oXMLHTTP.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write oXMLHTTP.responseBody
        oStream.SaveToFile vLocalFile, 2         ' 1 = no overwrite, 2 = overwrite
        oStream.Close
    End If

    'Clear memory
    Set oXMLHTTP = Nothing
End Function
```