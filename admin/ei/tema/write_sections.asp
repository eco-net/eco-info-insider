<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="shared/common.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim tema_id
tema_id=request("id")

Dim dgsheader,okheader,dgbheader,opslheader,newsheader,pagetext
dgsheader="<tr><td colspan='2'><span class='homeHeader'>De Gr&oslash;nne Sider </span>" &_
		"<span class='footer'> - det &oslash;kologiske modstykke til De Gule Sider</span><br>" &_
		"<table border='0' cellpadding='0' cellspacing='3' align='center' width='98%'>"
okheader="<tr><td colspan='2'><span class='homeHeader'>&Oslash;ko-Kalenderen</span>" &_ 
		"<span class='footer'> - guide til gr&oslash;nne events, kurser mv.</span><br>" &_
		"<table border='0' cellpadding='0' cellspacing='3' align='center' width='98%'>"&_
		"<tr><td><img src='/shared/graphics/layout/empty.gif'></td><td><img src=""/shared/graphics/layout/alertarrow.gif"" width=""14"" height=""8"">" &_
		"<a href=""/kurser/index.asp"">Tjek også <b>Kursuskalenderen</b></a></td></tr>"
dgbheader="<tr><td colspan='2'><span class='homeHeader'>Det Gr&oslash;nne Bibliotek</span>" &_ 
		"<span class='footer'> - &oslash;kologisk videns-samling</span><br>" &_
		"<table border='0' cellpadding='0' cellspacing='3' align='center' width='98%'>"&_
		"<tr><td><img src='/shared/graphics/layout/empty.gif'></td><td><img src=""/shared/graphics/layout/alertarrow.gif"" width=""14"" height=""8"">" &_
		"<a href=""/art/index.asp"">Se også <b>Artikeldatabasen</b></a></td></tr>"
opslheader="<tr><td colspan='2'><span class='homeHeader'>Opslagstavlen</span>" &_ 
		"<span class='footer'> - fra bruger til bruger</span><br>" &_
		"<table border='0' cellpadding='0' cellspacing='3' align='center' width='98%'>"
newsheader="<tr><td colspan='2'><span class='homeHeader'>Aktuelt</span>" &_ 
		"<span class='footer'> - Nyheder og andet aktualitetsstof</span><br>" &_
		"<table border='0' cellpadding='0' cellspacing='3' align='center' width='98%'>"
		
pagetext="<table border='0' cellpadding=4 cellspacing=0 align='center'>" 
		
set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_ecoinfo_STRING
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 3

Function getSek()
rs.Open()
while not rs.EOF
pagetext=pagetext & "<tr><td>"
if rs("tema_mrk")=1 then
pagetext=pagetext & "<img src='/shared/graphics/layout/tema.gif'></td><td>"
elseif rs("nyt_mrk")=1 then
pagetext=pagetext & "<img src='/shared/graphics/layout/nyt.gif'></td><td>"
else
pagetext=pagetext & "<img src='/shared/graphics/layout/empty.gif'></td><td>"
end if 
if rs("linktekst")<>"" then
	pagetext=pagetext & "<a href='" & rs("linktekst") & "'"
		if rs("sponsor")<>"" then
			pagetext=pagetext & " class='contentHeader2'"
		end if	
	pagetext=pagetext & "'>"& rs("tekst") & "</a>"
else
	pagetext=pagetext & rs("tekst")
end if 
pagetext=pagetext & "</td></tr>"
rs.MoveNext
wend
pagetext=pagetext & "</tr></table>"
rs.Close()
End Function

Function writeSections(choose)
pagetext=pagetext & dgsheader
rs.Source = "SELECT *  FROM eitema_sektion WHERE tema_id=" & tema_id & " AND sek='dgs' AND show=1 ORDER BY linie" 
getSek()
pagetext=pagetext & okheader
rs.Source = "SELECT *  FROM eitema_sektion WHERE tema_id=" & tema_id & " AND sek='ok' AND show=1 ORDER BY linie" 
getSek()
pagetext=pagetext & dgbheader
rs.Source = "SELECT *  FROM eitema_sektion WHERE tema_id=" & tema_id & " AND sek='dgb' AND show=1 ORDER BY linie" 
getSek()
pagetext=pagetext & opslheader
rs.Source = "SELECT *  FROM eitema_sektion WHERE tema_id=" & tema_id & " AND sek='opsl' AND show=1 ORDER BY linie" 
getSek()
pagetext=pagetext & newsheader
rs.Source = "SELECT *  FROM eitema_sektion WHERE tema_id=" & tema_id & " AND sek='news' AND show=1 ORDER BY linie" 
getSek()

pagetext=pagetext & "</table>" 
if choose = 0 then '**** temasektioner oprettes ****
filename="/log/ei/home/tema/inc_sections_" & tema_id & ".txt"
elseif choose = 1 then '**** tema vælges ****
filename="/log/ei/home/inc_sections.txt"
end if
WriteFile pagetext,filename
writeTemaLinks() '**** links på temaside ****
end function

Function writeTemaLinks()
listtext=""
rs.Source = "SELECT * FROM eitema_sektion WHERE tema_mrk=1 AND tema_id=" & tema_id & " ORDER BY sek"
rs.Open()
while not rs.EOF
if rs("linktekst")<>"" then
	listtext=listtext & "<a href='" & rs("linktekst") & "'"
		if rs("sponsor")<>"" then
			listtext=listtext & " class='contentHeader2'"
		end if	
	listtext=listtext & "'>"& rs("tekst") & "</a><br><br>"
end if
rs.MoveNext
wend
rs.Close()

filename="/log/ei/home/inc_temalinks.txt"
WriteFile listtext,filename
end Function

%>
