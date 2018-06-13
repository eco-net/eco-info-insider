<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<%
theSQL="UPDATE eiart_maindata SET authordate='' WHERE id=16"
execCommand theSQL
%>