```vb
Public Sub CheckForUpdates()
    
    'get the path of the current file
    Const RELEASE_URL As String = "https://api.github.com/repos/byronwall/butl/releases"
    
    Dim githubData As String
    githubData = DownloadFileAsString(RELEASE_URL)
    
    'this will grab the first file from the most recent release
    'this is a cheap way to "parse" the JSON without a library
    Dim splitURL As String
    splitURL = Split(Split(githubData, "tag_name"":")(1), """")(1)
    
    Dim currentVersion As String
    currentVersion = "Current version on GitHub is " & vbCrLf & _
                 vbTab & splitURL & vbCrLf & _
                 "Version of bUTL on computer is" & vbCrLf & _
                 vbTab & bUTL_GetVersion() & vbCrLf & _
                 "Do you want to update?"
    
    Dim shouldUpdate As VbMsgBoxResult
    shouldUpdate = MsgBox(currentVersion, vbYesNo, "Update?")
    
    If shouldUpdate = vbYes Then UpdateSelf


End Sub
```