<!--#include virtual="/shared/common.asp"-->
<!--#include virtual="/connections/ecoinfo.asp"-->

<%
DIM startText, modelText, endText, theSQL, theFile

'1. Creating dropdown/List options for amter
startText="document.write('<option value=""0"">&lt;V&aelig;lg&gt;</option>');" & chr(13)
modelText="document.write('<option value=""#0#"">#1#</option>');" & chr(13)
endText=""
theSQL="SELECT nr,navn FROM eisys_amter ORDER BY navn"
theFile="/log/optionfiles/amter_options.js"
Makefile theSQL,MM_ecoinfo_STRING,startText,endText,modelText,theFile

'1. Creating dropdown/List options for kommuner
startText="document.write('<option value=""0"">&lt;V&aelig;lg&gt;</option>');" & chr(13)
modelText="document.write('<option value=""#0#"">#1#</option>');" & chr(13)
endText=""
theSQL="SELECT nr,navn FROM eisys_kommuner ORDER BY navn"
theFile="/log/optionfiles/kommuner_options.js"
Makefile theSQL,MM_ecoinfo_STRING,startText,endText,modelText,theFile
%>