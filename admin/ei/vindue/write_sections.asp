<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/shared/common.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim tema_id
tema_id=request("id")

Dim dgsheader,okheader,dgbheader,opslheader,newsheader,pagetext
dgsheader="<tr><td colspan='2'><span class='homeHeader'><a href='/dgs/index.asp'>De Gr&oslash;nne Sider</a> </span>" &_
		"<span class='footer'> - det &oslash;kologiske modstykke til De Gule Sider</span><br>" &_
		"<table border='0' cellpadding='0' cellspacing='3' align='center' width='98%'>"
okheader="<tr><td colspan='2'><span class='homeHeader'><a href='/ok/index.asp'>&Oslash;ko-Kalenderen</a></span>" &_ 
		"<span class='footer'> - guide til gr&oslash;nne events, kurser mv.</span><br>" &_
		"<table border='0' cellpadding='0' cellspacing='3' align='center' width='98%'>"&_
		"<tr><td><img src='/shared/graphics/layout/empty.gif'></td><td><img src=""/shared/graphics/layout/alertarrow.gif"" width=""14"" height=""8"">" &_
		"<a href=""/kurser/index.asp"">Tjek også <b>Kursuskalenderen</b></a></td></tr>"
dgbheader="<tr><td colspan='2'><span class='homeHeader'><a href='/dgb/index.asp'>Det Gr&oslash;nne Bibliotek</a></span>" &_ 
		"<span class='footer'> - &oslash;kologisk videns-samling</span><br>" &_
		"<table border='0' cellpadding='0' cellspacing='3' align='center' width='98%'>"&_
		"<tr><td><img src='/shared/graphics/layout/empty.gif'></td><td><img src=""/shared/graphics/layout/alertarrow.gif"" width=""14"" height=""8"">" &_
		"<a href=""/art/index.asp"">Se også <b>Artikeldatabasen</b></a></td></tr>"
opslheader="<tr><td colspan='2'><span class='homeHeader'><a href='/opsl/index.asp'>Opslagstavlen</a></span>" &_ 
		"<span class='footer'> - fra bruger til bruger</span><br>" &_
		"<table border='0' cellpadding='0' cellspacing='3' align='center' width='98%'>"
newsheader="<tr><td colspan='2'><span class='homeHeader'><a href='/news/index.asp'>Aktuelt</a></span>" &_ 
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
'if rs("tema_mrk")=1 then
'pagetext=pagetext & "<img src='/shared/graphics/layout/tema.gif'></td><td>"
'elseif rs("nyt_mrk")=1 then
'pagetext=pagetext & "<img src='/shared/graphics/layout/nyt.gif'></td><td>"
'else
pagetext=pagetext & "<img src='/shared/graphics/layout/empty.gif'></td><td>"
'end if 
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
rs.Source = "SELECT *  FROM eivindue_sektion WHERE sek='dgs' ORDER BY linie" 
getSek()
pagetext=pagetext & okheader
rs.Source = "SELECT *  FROM eivindue_sektion WHERE sek='ok' ORDER BY linie" 
getSek()
pagetext=pagetext & dgbheader
rs.Source = "SELECT *  FROM eivindue_sektion WHERE sek='dgb' ORDER BY linie" 
getSek()
pagetext=pagetext & opslheader
rs.Source = "SELECT *  FROM eivindue_sektion WHERE sek='opsl' ORDER BY linie" 
getSek()
pagetext=pagetext & newsheader
rs.Source = "SELECT *  FROM eivindue_sektion WHERE sek='news' ORDER BY linie" 
getSek()

pagetext=pagetext & "</table>" 

filename="/log/ei/home/inc_sections.txt"
'*********************SKRIVNING TIL TEXTFIL UDKOMMENTERET: ********
'WriteFile pagetext,filename
'writeTemaLinks() '**** links på temaside ****
'*******************************************************************************
end function

Function writeTemaLinks()
listtext=""
rs.Source = "SELECT * FROM eivindue_sektion ORDER BY sek"
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
