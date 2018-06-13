<!--#include virtual="/shared/common.asp"-->
<!--#include virtual="/connections/ecoinfo.asp"-->

<%
DIM startText, modelText, endText, theSQL, theFile

'1. Creating dropdown/List options
startText="document.write('<option value=""0"" selected>&lt;V&aelig;lg&gt;</option>');" & chr(13)
modelText="document.write('<option value=""#0#"">#1#</option>');" & chr(13)
endText=""
theSQL="SELECT id,name FROM eical_cat_maindata WHERE siteid=1 ORDER BY name"
theFile="/log/optionfiles/okcat_options.js"
Makefile theSQL,MM_ecoinfo_STRING,startText,endText,modelText,theFile

'2. Creating sub-menu
startText="<nobr><a href=""list.asp?sektion=ok&oktime=30"">Denne måned</a></nobr> | " &_
	"<nobr><a href=""list.asp?sektion=ok&oktime=40"">Næste måned</a></nobr> | "
modelText="<nobr><a href=""list.asp?sektion=ok&okcat=#0#&okcatname=#1#"">#1#</a></nobr> | "
endText=""
theSQL="SELECT id,name FROM eical_cat_maindata WHERE siteid=1 ORDER BY name"
theFile="/shared/submenues/inc_ok.txt"
Makefile theSQL,MM_ecoinfo_STRING,startText,endText,modelText,theFile

'3. Creating Time-menu
DIM theMonth, theYear, theDate, i
startText="document.write('<option value=""0"" selected>&lt;V&aelig;lg&gt;</option>');" & chr(13) &_
	"document.write('<option value=""10"">Kommende arrangementer</option>');" & chr(13) &_
	"document.write('<option value=""20"">Afholdte arrangementer</option>');" & chr(13)

modelText="document.write('<option value=""#0#"">#1#</option>');" & chr(13)
endText=""
Filename="/log/optionfiles/oktime_options.js"

theFile=StartText
theMonth=DatePart("m",Date())
theYear=DatePart("yyyy",Date())
theDate=theYear & "-" & right("0" & CSTR(theMonth),2) & "-01"
For i=0 TO 12
	SELECT CASE theMonth
		CASE 1	theMonthname="Januar"
		CASE 2	theMonthname="Februar"
		CASE 3	theMonthname="Marts"
		CASE 4	theMonthname="April"
		CASE 5	theMonthname="Maj"
		CASE 6	theMonthname="Juni"
		CASE 7	theMonthname="Juli"
		CASE 8	theMonthname="August"
		CASE 9	theMonthname="September"
		CASE 10	theMonthname="Oktober"
		CASE 11	theMonthname="November"
		CASE 12	theMonthname="December"
	END SELECT
	theFile=theFile & replace(replace(modelText,"#0#",theDate),"#1#",theMonthName & " " & theYear)
	theMonth=theMonth+1
	IF theMonth=13 THEN theMonth=1
	IF theMonth=1 THEN theYear=theYear+1
	theDate=theYear & "-" & right("0" & CSTR(theMonth),2) & "-01"
NEXT

theFile=theFile & endText
WriteFile theFile,Filename

%>