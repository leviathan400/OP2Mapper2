Attribute VB_Name = "modUpdateChecker"
' Update Checker Function
'
' Usage: Call CheckForUpdate("0001", "https://software.outpostuniverse.org/myapp.htm", "https://github.com/leviathan400/myapp/releases")

Public Sub CheckForUpdate(ByVal CurrentBuild As String, ByVal VersionURL As String, ByVal DownloadURL As String)
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim serverVersion As String
    
    ' Fetch version from server (ServerXMLHTTP supports modern TLS/HTTPS)
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setOption 2, 13056  ' SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS (optional, for self-signed certs)
    http.Open "GET", VersionURL, False
    http.send
    
    If http.Status = 200 Then
        serverVersion = Trim(http.responseText)
        
        ' MsgBox (serverVersion)        ' Show version from server
        
        ' Compare versions - if server version is newer
        If StrComp(serverVersion, CurrentBuild, vbTextCompare) > 0 Then
            If MsgBox("A new version is available." & vbCrLf & vbCrLf & _
                      "Would you like to download it now?", vbYesNo + vbQuestion, "Update Available") = vbYes Then
                ' Open download URL in default browser
                Shell "rundll32.exe url.dll,FileProtocolHandler " & DownloadURL, vbNormalFocus
            End If
        End If
    End If
    
    Set http = Nothing
    Exit Sub
    
ErrorHandler:
    ' Silently fail - don't bother user if update check fails
    Set http = Nothing
End Sub


