<!--#include file="lib.asp"-->
<%
 doSave()
 
 response.redirect "artikel.asp?id=" & request("id") & "&name=" & request("name") & "&kapid=" & request("kapid")
%>