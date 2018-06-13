<!--#include virtual="/Connections/ecoinfo.asp" -->
<%Dim ID,size,theSize,theImg

Function getImage(ID,size)
'ID=request("id")
'size=request("size")

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
strSQL = "SELECT * FROM images where ID=" & ID & " AND " & theSize & "=1"

conn.open MM_ecoinfo_STRING
rs.Open strSQL, conn, 1
if not rs.EOF then
	theImg = "http://insider.eco-info.dk/admin/ei/billedbasen/" & theFolder & "/" & rs("filename")
	else
theImg="http://insider.eco-info.dk/shared/graphics/spacer.gif"
end if
rs.close
conn.close
getImage=theImg
end function

%>
