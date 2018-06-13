<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<% 
Dim update 
if request("submit2")<>"" then
update=0
end if
if request("submit3")<>"" then
update=1
end if

%>


<%

theSQL="UPDATE eisys_insiderusers SET nomail=" & update & " WHERE orgid=" & request("id") 
			execCommand(theSQL)
			
%>
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
<center class="listheader">
<%
response.write "Organisationen med id: " & request("id") & "<br>"
			if update=1 then 
			response.write(" er afmeldt insidermail")
			else
			response.write(" er tilmeldt insidermail")
			end if
%>
</center>
<form name="form1" method="post" action="">
<div align="center">
<input type="button" name="Button" value="OK" class="formSubmitbutton" onClick="history.go(-1);">
</div>
</form>


