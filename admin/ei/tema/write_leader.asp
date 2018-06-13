<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="shared/common.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim tema_id
tema_id=request("id")
'tema_id="1"

Dim pagetext		
pagetext="<table width='388' border='0' cellspacing='0' cellpadding='0'><tr><td>"

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_ecoinfo_STRING
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 3

Function writeLeader(choose)
filnavn=""
pagetext=pagetext & dgsheader
rs.Source = "SELECT *  FROM eitema_maindata WHERE id=" & tema_id 
rs.Open()
if rs("img")>0 then
img_id=rs("img")
branding="B"
img=getFilename(img_id,branding)
filnavn="http://insider.eco-info.dk/admin/ei/billedbasen/branding/" & img
end if
pagetext=pagetext & "<img src='" & filnavn & "' border='0'>"

pagetext=pagetext & "</td><td valign='top' class='homeBigbox'>"
pagetext=pagetext & "<table width='100%' border='0' cellspacing='0' cellpadding='7'>"
pagetext=pagetext & "<tr><td class='home'><span class='homeHeader'>" & rs("title") & "</span><br><br>"
pagetext=pagetext & rs("descr")
if rs("link")<>"" then
pagetext=pagetext & "<br><br><a href='temaside.asp'>" & rs("linktekst") & "</a>"
end if
pagetext=pagetext & "</td></tr></table></td></tr><tr>"
pagetext=pagetext & "<td background='/shared/graphics/layout/dots_horz.gif' height='1' colspan='3'>"
pagetext=pagetext & "<img src='/shared/graphics/spacer.gif' width='3' height='1'></td></tr></table>"

rs.Close()

if choose = 0 then
filename="/log/ei/home/tema/inc_leader_" & tema_id & ".txt"

elseif choose = 1 then
filename="/log/ei/home/inc_leader.txt"
end if
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
