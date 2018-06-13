<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim rsMail__MMColParam
rsMail__MMColParam = ""
if (Request("mailadr") <> "") then rsMail__MMColParam = Request("mailadr")
%>
<%
set rsMail = Server.CreateObject("ADODB.Recordset")
rsMail.ActiveConnection = MM_ecoinfo_STRING
rsMail.Source = "SELECT eiorg_maindata.id, eiorg_maindata.name, eisys_insiderusers.username,eisys_insiderusers.password, eisys_insiderusers.email  FROM eiorg_maindata INNER JOIN eisys_insiderusers ON eiorg_maindata.id = eisys_insiderusers.orgid  WHERE email = '" + Replace(rsMail__MMColParam, "'", "''") + "';"
rsMail.CursorType = 0
rsMail.CursorLocation = 2
rsMail.LockType = 3
rsMail.Open()
rsMail_numRows = 0
%>
<% if rsMail.EOF then
response.write("Findes ikke")
response.end
else
%>
<%
Dim theTo,theFrom,theSubject,theBody
theTo = rsMail__MMColParam
theFrom="eco-net@eco-net.dk"
theSubject="Brugerdata til Øko-info"
theBody="Som svar på forespørgsel om glemte brugerdata er her -- " 
if rsMail__MMColParam <>"" then
	if rsMail.Fields.Item("name").Value<>"" then
		theBody=theBody & rsMail.Fields.Item("name").Value & "'s"
	end if
end if
theBody=theBody & " --Brugernavn: " & rsMail.Fields.Item("username").Value 
theBody=theBody & "--Adgangskode: " & rsMail.Fields.Item("password").Value
theBody=theBody & "--Med Venlig hilsen--Øko-net--"
theBody=theBody & "<a href='http://www.eco-info.dk'>eco-info.dk</a>"
SendCDOMail theTo,theFrom,theSubject,theBody

end if
%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p>Vi sender straks en mail med dine brugerdata </p>
<p>Med venlig hilsen </p>
<p>&Oslash;ko-net<br>
62244324<br>
</p>
</body>
</html>
<%
rsMail.Close()
%>
<% if rsOrg__MMColParam <>"" and rsOrg__MMColParam <> "0" then
rsOrg.Close()
end if
%>
