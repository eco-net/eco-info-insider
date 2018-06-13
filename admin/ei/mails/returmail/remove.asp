<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp"-->
<%

server.ScriptTimeout = 3000000 
delnumbers=10
if request("numbers_del")<>"" then
delnumbers=CInt(request("numbers_del"))
end if
set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_ecoinfo_STRING
rs.Source = "SELECT * FROM eisys_returnmail"
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 3
rs.Open()
response.write "OK"
response.end
i=0
while not rs.EOF and i<delnumbers
i=i+1
returnmail=returnmail & rs("mail") & ";" ' returnerede mailadresser
rs.MoveNext
wend
rs.Close
fmpmail=""
fmpantal_start=antalFmp()' skriver også fmpmail strengen

fmpArray=Split(fmpmail,";")'fmp til array
returnArray=Split(returnmail,";")'retur til array
returantal=Ubound(returnArray)

if returantal<delnumbers then
delnumbers=returantal
end if
for i=0 to delnumbers-1
inFmp returnArray(i)'sletter også
'response.write returnArray(i) & delnumbers & "<br>"
next

fmpantal_slut=antalFmp()

response.write "mailadresser i eisys_fmpmail før kørsel: " & fmpantal_start & "<br>"
response.write "mailadresser der er slettet fra eisys_returnmail og eisys_fmpmail: " & returantal & "<br>"
response.write "der er stadig: " & antalRetur() & " mailadresser der skal slettes<br>"
response.write "mailadresser i eisys_fmpmail efter kørsel: " & fmpantal_slut & "<br>"

Function inFmp(m)
for j=0 to fmpantal_start
if fmpArray(j)=m then
del(m)

end if
'response.write m & fmpArray(j) & "<br>"
next
end function

Function antalFmp()
rs.Source = "SELECT * FROM eisys_fmpmail"
rs.Open
a=0
if not rs.EOF then
while not rs.EOF
a=a+1
fmpmail=fmpmail & rs("mail") & ";" ' fmp mailadresser

rs.MoveNext
wend
end if
rs.Close
antalFmp=a
end function

Function antalRetur()
rs.Source = "SELECT * FROM eisys_returnmail"
rs.Open
b=0
if not rs.EOF then
while not rs.EOF
b=b+1
rs.MoveNext
wend
end if
rs.Close
antalRetur=b
end function

Function del(rm)
theSQL="DELETE FROM eisys_fmpmail WHERE mail='" & rm & "'"
'response.write theSQL & "<br>"
execCommand theSQL
theSQL="DELETE FROM eisys_returnmail WHERE mail='" & rm & "'"
'response.write theSQL & "<br>"
execCommand theSQL
end function
%>
<html><body>
<p>Slet mailadresser der ikke virker.</p>
<p>Tag lidt ad gangen.  </p>
<form name="form1" method="post" action="">
  <label>
  <input name="numbers_del" type="text" id="numbers_del" value="10">
  </label>
  <label>
  <input type="submit" name="Submit" value="Submit">
  </label>
</form>
</body></html>
