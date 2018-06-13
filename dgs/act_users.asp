<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->

<%
userpw=GetUniquePW(request("username"),request("password"))
theSQL="UPDATE eisys_insiderusers SET username='" & request("username") & "'," &_
	"password='" & userpw & "'," &_
	"email='" & request("email") & "' " &_
	"WHERE orgid=" & request("orgid")
execCommand theSQL
response.redirect ("users.asp?done=1")
%>