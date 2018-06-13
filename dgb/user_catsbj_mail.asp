<!--#include virtual="/shared/common.asp" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%Dim Subj,Body,theTo,theFrom,theSubject
theTo="nhl@eco-net.dk"
theFrom="eco-net@eco-net.dk"
theSubject="Ændring af Emneord for Publikation"

Subj=""

FOR each Item in request("selSbj")
Subj=Subj & Item & ", "
next 
Body="Publikationen: "  & request("libname") & " med id nr: " & request("id") 
Body=Body & " skal have rettet emneord til: " & Subj

SendMail theTo,theFrom,theSubject,Body
%>
Øko-net har modtaget dine ønsker om rettelser af Emneord.<BR>
Vi retter snarest muligt.<BR>
MVH Øko-net 62244324.

</body>
</html>
