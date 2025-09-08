<%@ Language="VBScript" %>
<%
Option Explicit

Dim uploadDir
uploadDir = Server.MapPath(".") & "\uploads\"

Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(uploadDir) Then
    fso.CreateFolder(uploadDir)
End If

Dim totalBytes, dataBin
totalBytes = Request.TotalBytes

If totalBytes > 0 Then
    dataBin = Request.BinaryRead(totalBytes)

    Dim fileName, filePath, stream
    fileName = "upload_" & Replace(Replace(CStr(Timer), ".", "_"), ":", "_") & ".bin"
    filePath = uploadDir & fileName

    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write dataBin
    stream.SaveToFile filePath, 2 ' Overwrite
    stream.Close

    Response.Write "Saved upload as: " & fileName
Else
    Response.Write "No data received"
End If

Set stream = Nothing
Set fso = Nothing
%>
