<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim rsPageData__theID
rsPageData__theID = "0"
if (request("id") <> "") then rsPageData__theID = request("id")

Set Upload = Server.CreateObject("Persits.Upload.1")

' Upload files
'Upload.OverwriteFiles = False ' Generate unique names
'Upload.SetMaxSize 100000,true ' Truncate files above 1MB
'Upload.Save "D:\www2\eco-info\insider.oko-info\news\images\" 
Count=Upload.SaveVirtual("images/")

Set File = Upload.Files("file")
If Not File Is Nothing Then
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_ecoinfo_STRING
	'rs.Source = "SELECT *  FROM bu_Artikel WHERE bu_Artikel.Artikel_ID = " + Replace(rsPageData__theID, "'", "''") + ""
	rs.Source = "SELECT *  FROM bu_Artikel WHERE bu_Artikel.Artikel_ID = 687"
	rs.CursorLocation = 2
	rs.LockType = 3
	rs.Open()
	rs("theImage") = File.Binary
	rs.Update
	rs.close
		Response.Write "File saved."
	Else
		Response.Write "File not selected."
End If

%> 

