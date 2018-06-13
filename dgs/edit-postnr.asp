<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<%
'theSQL="UPDATE eisys_postnr SET provins='False',land=1 WHERE postnr=0009"
theSQL="Insert into eisys_postnr (postnr,city,provins,land) Values (0999,'København C','False',1)"
execCommand theSQL
%>