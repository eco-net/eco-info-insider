
<!--#include virtual="/shared/common.asp"-->
<%
id=request("id")
userid=request("userid")
name=request("name")
firstname=request("firstname")
lastname=request("lastname")
theto=request("email")
username=request("username")
password=request("password")
password=GetUniquePW(username,password)
theSQL="UPDATE eisys_insiderusers SET username='" & username & "',password='" & password & "' WHERE id=" & userid
execCommand theSQL

thetext="<p><B>Dine Brugeroplysninger til &Oslash;ko-info<B></p><p><a href='http://www.eco-info.dk'>www.eco-info.dk</a></p><p>Brugernavn: " & username & "</p><p>Agangskode: " & password & "</p><p>MVH &Oslash;ko-net<br>62 24 43 24<br><a href='mailto:kmh@eco-net.dk'>kmh@eco-net.dk</a></p>"
theBody="<html><head>" & Style & "</head><body>" & thetext & "</body></html>"
SendCDOHtmlMail theto,"eco-net@eco-net.dk","Brugeroplysninger rettet på Øko-info",theBody

response.redirect "edit.asp?id=" & id
%>