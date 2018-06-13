<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp"-->
<%
antal=request("members")
theSQL="INSERT INTO eisys_newmember (members) Values(" & antal & ")"
'response.write theSQL
'response.end
execCommand theSQL
response.redirect "index.asp"
%>