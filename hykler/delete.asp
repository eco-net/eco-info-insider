<%
strConnect = "PROVIDER=MSDASQL;DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("db.mdb") & ";"
theSQL="DELETE FROM tekst"
DIM theComm
Set theComm = Server.CreateObject("ADODB.Command")
theComm.ActiveConnection = strConnect
theComm.CommandText = theSQL
theComm.Execute
theComm.ActiveConnection.Close
theComm=""
response.redirect "index.asp"
%>
