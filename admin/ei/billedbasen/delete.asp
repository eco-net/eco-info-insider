<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim rsImg__theID,filename
rsImg__theID = "0"
if (request("id") <> "") then rsImg__theID = request("id")
%>
<%
set rsImg = Server.CreateObject("ADODB.Recordset")
rsImg.ActiveConnection = MM_ecoinfo_STRING
rsImg.Source = "SELECT *  FROM images  WHERE id = " + Replace(rsImg__theID, "'", "''") + ""
rsImg.CursorType = 1
rsImg.CursorLocation = 2
rsImg.LockType = 3
rsImg.Open()
rsImg_numRows = 0
%>
<%if not rsImg.EOF then
Set fso = CreateObject("Scripting.FileSystemObject")
filename=rsImg("filename")
if rsImg("twocol")=1 then
del("2col")
end if
if rsImg("twocol")=1 then
del("2col")
end if
if rsImg("threecol")=1 then
del("3col")
end if
if rsImg("tema")=1 then
del("tema")
end if
if rsImg("rightcol")=1 then
del("rightcol")
end if
if rsImg("thumbnail")=1 then
del("thumbnail")
end if
if rsImg("branding")=1 then
del("branding")
end if
if rsImg("splash")=1 then
del("splash")
end if
del("upload")

rsImg.Close()
theSQL="Delete from images WHERE id=" & rsImg__theID & ";" & chr(13)
execCommand theSQL

response.redirect "index.asp"
else
response.write "billedet er ikke registreret i DB"
end if
Function del(folder)
if (fso.FileExists (Server.MapPath(folder & "/")  & "/" & filename)) then
fso.DeleteFile (Server.MapPath(folder & "/")  & "/" & filename)
end if
end function
%>