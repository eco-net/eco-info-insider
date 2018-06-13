<!--#include virtual="shared/common.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim tema_id
tema_id=request("id")
'tema_id="1"

Dim pagetext		
pagetext="<table width='100%' height='100%' border='0' cellspacing='0' cellpadding='0'>"
pagetext=pagetext & "<tr><td valign='top'><div align='center'><span class='contentHeader1'><br>"

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_ecoinfo_STRING
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 3

Function writeArticle(choose)
rs.Source = "SELECT *  FROM eitema_maindata WHERE id=" & tema_id
rs.Open()
header=rs("header")
subheader=rs("subheader")
rs.Close()
pagetext=pagetext & header & "</span><br>"
pagetext=pagetext & subheader & "</div></td></tr><tr><td valign='top'><p>&nbsp;</p>"

rs.Source = "SELECT *  FROM eitema_artikel WHERE tema_id=" & tema_id & " ORDER BY place"
rs.Open()


pagetext=pagetext & "<img src='/admin/ei/shared/showimg.asp?id=" & rs("img") & "' border='0'>"

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
%>
