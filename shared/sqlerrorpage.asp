<!--#include file="common.asp"-->
<%
theBody="En bruger er blevet afvist fra siden: " & request("p") & "<br>med brug af dette blacklistede ord: " & request("e")
SendCDOHtmlMail "web@eco-net.dk","insider@eco-net.dk","SQLERROR",theBody
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Fejl</title>
</head>

<body>
<h1 align="center">Fejl</h1>
<p align="center">Ordet: <%=request("e")%> er ikke tilladt.</p>
<p align="center"><a href="javascript:history.go(-1)">Tilbage</a></p>
</body>
</html>
