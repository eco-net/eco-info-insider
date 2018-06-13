<!--#include virtual="asptest.asp"-->
<!--#include virtual="/connections/dbadmin.asp"-->
<%
MakeFile "Select id,code,logtime from sqlstatements ORDER by id",MM_dbadmin_STRING,"<table border=""1"">","</table>","<tr><td>#0#</td><td>#1#</td><td>#2#</td></tr>","data.txt"
response.write "<br>OK"
%>