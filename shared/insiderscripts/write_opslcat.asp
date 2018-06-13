<!--#include virtual="/shared/common.asp"-->
<!--#include virtual="/connections/ecoinfo.asp"-->

<%
DIM startText, modelText, endText, theSQL, theFile

'1. Creating dropdown/List options
startText="document.write('<option value=""0"" selected>&lt;V&aelig;lg&gt;</option>');" & chr(13)
modelText="document.write('<option value=""#0#"">#1#</option>');" & chr(13)
endText=""
theSQL="SELECT id,name FROM eiopsl_cat_maindata ORDER BY name"
theFile="/log/optionfiles/opslcat_options.js"
Makefile theSQL,MM_ecoinfo_STRING,startText,endText,modelText,theFile

'2. Creating sub-menu
startText=""
modelText="<nobr><a href=""list.asp?sektion=dgb&dgbcat=#0#&dgbcatname=#1#"">#1#</a></nobr> | "
endText=""
theSQL="SELECT id,name FROM eiopsl_cat ORDER BY name"
theFile="/shared/submenues/inc_opsl.txt"
Makefile theSQL,MM_ecoinfo_STRING,startText,endText,modelText,theFile
%>