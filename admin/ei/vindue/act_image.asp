<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/common.asp" -->
<%
if request("update")<>"" then
theSQL="UPDATE eivindue_item SET img_id=" & request("img_id") &_
",img_text='" & request("img_text") &_
"',col=" & request("radiobutton") &_
",thesort=" & request("sort") &_
"  WHERE id=" & request("item_id")
else
theSQL="INSERT INTO eivindue_item (img_id,img_text,col,thesort,page_id)" &_
" VALUES (" & request("img_id") & ",'" &_
request("img_text") & "'," &_
request("radiobutton") & "," &_
request("sort") & "," &_
request("page_id") & ")"
end if
'response.write theSQL
'response.end
execCommand theSQL
response.redirect "edit_artikler.asp?id=" & request("page_id") & "&name=" & request("name")
%>

