```vb
Function DownloadFileAsString(ByVal vWebFile As String) As String
    Dim oXMLHTTP As Object, i As Long, vFF As Long, oResp() As Byte

    'You can also set a ref. to Microsoft XML, and Dim oXMLHTTP as MSXML2.XMLHTTP
    Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    oXMLHTTP.Open "GET", vWebFile, False         'Open socket to get the website
    oXMLHTTP.Send                                'send request

    'Wait for request to finish
    Do While oXMLHTTP.readyState <> 4
        DoEvents
    Loop

    DownloadFileAsString = oXMLHTTP.responseText 'Returns the results as a byte array

    'Clear memory
    Set oXMLHTTP = Nothing
End Function
```