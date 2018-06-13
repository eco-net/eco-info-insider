<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/common.asp" -->
<%
if request("update")<>"" then
theSQL="UPDATE eivindue_item SET tekst='" & request("description") &_
"',source='" & request("source") &_
"',link='" & request("link") &_
"',linktext='" & request("linktext") &_
"',col=" & request("radiobutton") &_
",thesort=" & request("sort") &_
" WHERE id=" & request("item_id")
else
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
end if
execCommand theSQL
response.redirect "edit_artikler.asp?id=" & request("page_id") & "&name=" & request("name")
%>

