<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%Dim ID,size,theSize,theFolder
ID=request("id")
size=request("size")

if size="2" then
theSize="twocol"
theFolder="2col"
elseif size="3" then
theSize="threecol"
theFolder="3col"
elseif size="R" then
theSize="rightcol"
theFolder="rightcol"
elseif size="B" then
theSize="branding"
theFolder="branding"
elseif size="S" then
theSize="splash"
theFolder="splash"
elseif size="T" then
theSize="tema"
theFolder="tema"
end if
set conn=Server.CreateObject("ADODB.Connection")
set rs= Server.CreateObject("ADODB.Recordset")
if size<>"U" then
strSQL = "SELECT * FROM images where ID=" & ID & " AND " & theSize & "=1"
else
strSQL = "SELECT * FROM images where ID=" & ID
theFolder="upload"
end if

conn.open MM_ecoinfo_STRING
rs.Open strSQL, conn, 1
	response.write "<img src='" & theFolder & "/" & rs("filename") & "'>"
 
rs.close
conn.close
%>
