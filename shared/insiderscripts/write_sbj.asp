<!--#include virtual="/shared/common.asp"-->
<!--#include virtual="/connections/ecoinfo.asp"-->

<%
DIM startText, modelText, endText, theSQL, theFile

'1. Creating dropdown/List options
startText="document.write('<option value=""0"">&lt;V&aelig;lg&gt;</option>');" & chr(13)
modelText="document.write('<option value=""#0#"">#1#</option>');" & chr(13)
endText=""
theSQL="SELECT id,name FROM eisbj_maindata WHERE siteid=1 ORDER BY name"
theFile="/log/optionfiles/sbj_options.js"
Makefile theSQL,MM_ecoinfo_STRING,startText,endText,modelText,theFile
%>