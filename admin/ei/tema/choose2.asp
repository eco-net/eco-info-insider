<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->

<!--#include file="write_leader.asp"-->

<%
writeLeader(1)

response.redirect ("confirmation.asp?done=1")
%>