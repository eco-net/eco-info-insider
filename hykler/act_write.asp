<%
<!--#include virtual="/shared/sqlcheck.asp"-->
strConnect = "PROVIDER=MSDASQL;DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("db.mdb") & ";"
theSQL="INSERT INTO tekst (dato,titel,forfatter,tekst) VALUES ('" &_
Date() & "','" &_
request("titel") & "','" &_
request("forfatter") & "','" &_
replace(request("tekst"),vbCrLf,"<br><br>") & "')" 

DIM theComm
Set theComm = Server.CreateObject("ADODB.Command")
theComm.ActiveConnection = strConnect
theComm.CommandText = theSQL
theComm.Execute
theComm.ActiveConnection.Close
theComm=""
response.redirect "index.asp"

%>