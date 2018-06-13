<!--#include virtual="/Connections/ecoinfo.asp" -->
<%Dim ID
ID=request("id")
set conn=Server.CreateObject("ADODB.Connection")
set rs= Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM image_maindata where ID=" & ID

conn.open MM_ecoinfo_STRING
rs.Open strSQL, conn, 1
	strImage = rs("image").GetChunk(1024000)
	if rs("ext")="JPG" then
    Response.ContentType = "image/jpg"
	elseif rs("ext")="GIF" then
    Response.ContentType = "image/gif"
	elseif rs("ext")="BMP" then
    Response.ContentType = "image/bmp"
	end if 
    Response.BinaryWrite strImage
 
rs.close
conn.close
%>
