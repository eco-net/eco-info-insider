<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<%
theSQL="DELETE FROM eivindue_item WHERE id=" & request("item_id")
execCommand theSQL
response.redirect "edit_artikler.asp?id=" & request("page_id") & "&name=" & request("name")
%>