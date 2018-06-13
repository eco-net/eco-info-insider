<!--#include virtual="/shared/common.asp"-->
<!--#include virtual="/connections/ecoinfo.asp"-->

<%
DIM startText, modelText, endText, theSQL, theFile

'1. Creating dropdown/List options
startText="document.write('<option value=""0"" selected>&lt;V&aelig;lg&gt;</option>');" & chr(13)
modelText="document.write('<option value=""#0#"">#1#</option>');" & chr(13)
endText=""
theSQL="SELECT id,name FROM eiorg_cat_maindata WHERE siteid=1 ORDER BY name"
theFile="/log/optionfiles/dgscat_options.js"
Makefile theSQL,MM_ecoinfo_STRING,startText,endText,modelText,theFile

'2. Creating sub-menu
startText=""
modelText="<nobr><a href=""list.asp?sektion=dgs&dgscat=#0#&dgscatname=#1#"">#1#</a></nobr> | "
endText=""
theSQL="SELECT id,name FROM eiorg_cat_maindata WHERE siteid=1 ORDER BY name"
theFile="/shared/submenues/inc_dgs.txt"
Makefile theSQL,MM_ecoinfo_STRING,startText,endText,modelText,theFile
%>