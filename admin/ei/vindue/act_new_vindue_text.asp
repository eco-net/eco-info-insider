<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/common.asp" -->
<%
theSQL="INSERT INTO eivindue_item (tekst,source,link,linktext,col,thesort,page_id)" &_
" VALUES ('" & request("description") & "','" &_
request("source") & "','" &_
request("link") & "','" &_
request("linktext") & "'," &_
request("radiobutton") & "," &_
request("sort") & "," &_
request("page_id") & ")"
'response.write theSQL
'response.end
execCommand theSQL
response.redirect "edit_artikler.asp?id=" & request("page_id") & "&name=" & request("name")
%>

