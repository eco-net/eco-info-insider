<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp"-->
<%
theSQL="DELETE FROM eisys_returnmail"
execCommand theSQL
antal=0
fmpmail=""
fmpstr=""
set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_ecoinfo_STRING
rs.Source = "SELECT * FROM eisys_fmpmail"
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 3
rs.Open()
while not rs.EOF 
fmpmail=fmpmail & rs("mail") & ";" ' Lars udvalgte mailadresser
rs.MoveNext
wend
fmpArray=Split(fmpmail,";")
rs.Close
mailtext=""
theText=ReadFile("returmails.txt") ' mailadresser i mailbody fra de returnerede mails
mailArray=Split(theText,"#")
for i=0 to Ubound(mailArray)
mailArray(i)=Replace(mailArray(i),vbCr," ")
mailArray(i)=Replace(mailArray(i),VbCrLf," ")
mailArray(i)=Replace(mailArray(i),vbLf," ")
mailArray(i)=Replace(mailArray(i),vbNewLine," ")
mailArray(i)=Replace(mailArray(i),vbTab," ")
Words=Split(mailArray(i)," ") ' en mailadresse ad gangen
for j=0 to Ubound(Words)
if InStr(Words(j),"@")>0 And InStr(Words(j),".")>0 then
inFmp Words(j)
mailtext=mailtext & Words(j) & ";"
end if
next
next
response.write fmpstr ' udskrift af Lars mailadresser der kom retur
response.write antal & "<br>"
response.write "<a href='http://insider.eco-info.dk/dbAdmin/pages/struct_tables.asp'>Find tabellen eisys_returnmail </a>"
'WriteFile mailtext, "mailtext.txt"
Function inFmp(m) ' findes den returnerede mail i Lars udvalgte?
for ij=0 to UBound(fmpArray)
if InStr(fmpArray(ij),m)>0 then
inFmpstr fmpArray(ij)
end if
next
end function
Function inFmpstr(m2) ' registrering af Lars mailadresser der kom retur
if InStr(fmpstr,m2)=0 then
toDB m2
fmpstr=fmpstr & m2 & "<br>"
antal=antal+1
end if
end function
Function toDB(mail) 'til databasen
theSQL="INSERT INTO eisys_returnmail (mail) VALUES ('" & mail & "')"
execCommand theSQL
end function
%>