<!--#include virtual="/shared/common.asp" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%Dim Cat,Subj,Body,theTo,theFrom,theSubject
theTo="nhl@eco-net.dk"
theFrom="eco-net@eco-net.dk"
theSubject="Ændring af kategori og Emneord for Arrangement"
Cat=""
Subj=""
FOR each Item in request("selCat")
Cat=Cat & Item & ", "
next 
FOR each Item in request("selSbj")
Subj=Subj & Item & ", "
next 
Body="Arrangementet: "  & request("arrname") & " med id nr: " & request("id") 
Body=Body & " ønsker disse kategorier: " & Cat
Body=Body & " og disse emneord: " & Subj

SendMail theTo,theFrom,theSubject,Body
%>

Øko-net har modtaget dine ønsker om rettelser af Kategori/Emneord.<BR>
Vi retter snarest muligt.<BR>
MVH Øko-net 62244324.

</body>
</html>
