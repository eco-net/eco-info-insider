<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/shared/common.asp"-->
<!--include file="write_sections.asp"-->

<%
theSQL="UPDATE enblad_maindata SET chosen=0"
execCommand(theSQL)
theSQL="UPDATE enblad_maindata SET chosen=1 WHERE id=" & request("id")
execCommand(theSQL)
'writeSections(1)

response.redirect "index.asp?id=" & request("id")
%>