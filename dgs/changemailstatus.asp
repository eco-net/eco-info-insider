<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_access.asp" -->
<!--#include virtual="/shared/common.asp" -->
<% IF CINT(request("change"))=1 then ' sæt nomail=1 -> modtager ikke mails fremover
'	response.write "Modtager ikke<bR><br>"
	theSQL="Update eisys_insiderusers set nomail=1 where orgid="&request("orgid")
'	response.write theSQL
	execCommand theSQL
ELSEIF CINT(request("change"))=2 then 'sæt nomail=0 -> modtager mails fremover
'	response.write "Modtager<br><br>"
	theSQL="Update eisys_insiderusers set nomail=0 where orgid="&request("orgid")
'	response.write theSQL
	execCommand theSQL
END IF

set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT m.name,m.firstname,m.middlename,m.lastname,i.nomail FROM eiorg_maindata m  INNER JOIN eisys_insiderusers i ON m.id=i.orgid where m.id="&request("orgid")
rsPageData.open()
if len(rsPageData("name"))<3 then
	strName=rsPageData("firstname")&" "&rsPageData("lastname")
else
	strName=rsPageData("name")
end if
%>
<html>
<head>
<title>Skift Mailmodtagelses</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" action="" method="post">
<%
if CINT(rsPageData("nomail"))=1 then ' ønsker ikke at modtage mail
	response.write "<b>"&strName&"</b><br>modtager i øjeblikket <b>ikke</b> mails.<bR><br>Klik nedenfor, for at aktivere modtagelsen."
	response.write "<br><br><input type=""button"" value=""Aktiver mailmodtagelse"" onclick=""window.location='changemailstatus.asp?orgid="&request("orgid")&"&change=2'"">"
else ' vil gerne modtage mail
	response.write "<b>"&strName&"</b><br>modtager i øjeblikket mails.<bR><br>Klik nedenfor, for at deaktivere modtagelsen."
	response.write "<br><br><input type=""button"" value=""Deaktiver mailmodtagelse"" onclick=""window.location='changemailstatus.asp?orgid="&request("orgid")&"&change=1'"">"
end if 
%>
</form>
</body>
</html>
<%
set rsPageData=nothing
%>