<%@ Language="VBScript" %>
<%
Option Explicit

' Save directory — ensure this exists and IIS has write permissions
Dim uploadDir
uploadDir = Server.MapPath(".") & "\uploads\"

' Ensure upload folder exists
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(uploadDir) Then
    fso.CreateFolder uploadDir
End If

' Read method
Dim method
method = Request.ServerVariables("REQUEST_METHOD")

Response.Write "<h3>Request Method: " & method & "</h3>"

' Read raw bytes from the request
Dim totalBytes, dataBin
totalBytes = Request.TotalBytes
Response.Write "<p>TotalBytes: " & totalBytes & "</p>"

If totalBytes > 0 Then
    dataBin = Request.BinaryRead(totalBytes)
    
    ' Save to file (you could parse boundary if needed)
    Dim fileName, filePath, stream
    fileName = "upload_" & Replace(Replace(CStr(Timer), ".", "_"), ":", "_") & ".bin"
    filePath = uploadDir & fileName

    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write dataBin
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close

    Response.Write "<p>Saved raw request data to: " & fileName & "</p>"
Else
    Response.Write "<p>No data received.</p>"
End If

Set stream = Nothing
Set fso = Nothing
%>
