<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/shared/common.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%

Dim pagetext		
pagetext="<table width='388' border='0' cellspacing='0' cellpadding='0'>"

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_ecoinfo_STRING
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 3

Function writeLeader()
filnavn=""
pagetext=pagetext & dgsheader
rs.Source = "SELECT *  FROM en_mainpage"
rs.Open()
pagetext=pagetext & "<tr><td colspan='3' align='center'><b>" & rs("title") & "</b></td></tr>"
if rs("img")>0 then
img_id=rs("img")

img=getFilename(img_id,"3")
filnavn="http://insider.eco-info.dk/admin/ei/billedbasen/3col/" & img
end if
if rs("link")<>"" then
pagetext=pagetext & "<tr><td align='center' colspan='3'><a href='" & rs("link") & "' title='" & rs("linktext") & "'><img src='" & filnavn & "' border='0'></a></td></tr>"
else
pagetext=pagetext & "<tr><td align='center' colspan='3'><img src='" & filnavn & "' border='0'></td></tr>"
end if
pagetext=pagetext & "<tr><td width='20'>&nbsp;</td><td>" & rs("descr") & "</td>" 
if rs("link")<>"" then
pagetext=pagetext & "<td width='50'><a href='" & rs("link") & "'>" & rs("linktext") & "</a><br><br></td></tr>"
else
pagetext=pagetext & "<td>&nbsp;<br><br></td></tr>"
end if
pagetext=pagetext & "<tr><td background='/shared/graphics/layout/dots_horz.gif' height='1' colspan='3'>"
pagetext=pagetext & "<img src='/shared/graphics/spacer.gif' width='3' height='1'></td></tr></table>"

rs.Close()


filename="/log/econet/inc_leader.txt"

WriteFile pagetext,filename

end function
Function getFilename(ID,size)
if size="2" then
theSize="twocol"
elseif size="3" then
theSize="threecol"
elseif size="R" then
theSize="rightcol"
elseif size="B" then
theSize="branding"
elseif size="S" then
theSize="splash"
elseif size="T" then
theSize="tema"
end if

set conn=Server.CreateObject("ADODB.Connection")
set rs2= Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM images where ID=" & ID & " AND " & theSize & "=1"

conn.open MM_ecoinfo_STRING
rs2.Open strSQL, conn, 1
getFilename= rs2("filename")
rs2.Close()
conn.Close()
end function

%>
