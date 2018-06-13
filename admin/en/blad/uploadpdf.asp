<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="shared/common.asp"-->
<%

Set Upload = Server.CreateObject("Persits.Upload.1")

' Upload files
Upload.OverwriteFiles = False ' Generate unique names
'Upload.SetMaxSize 100000,true ' Truncate files above 1MB
Count=Upload.SaveVirtual("pdf/") 
Set File = Upload.Files("file")
If Not File Is Nothing Then
		Set rs = Server.CreateObject("adodb.recordset")
		theSQL="UPDATE enblad_maindata SET pdf='" & File.ExtractFileName & "' WHERE id=" & Upload.Form("id")
		execCommand theSQL
		Response.Write "Pdf filen er uploaded"
	Else
		Response.Write "File not selected."
End If
%> 

