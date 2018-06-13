<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<%
id=CInt(request("id"))
theSQL="DELETE FROM eiopsl_maindata WHERE id=" & id
execCommand theSQL
response.redirect "index.asp"
%>
