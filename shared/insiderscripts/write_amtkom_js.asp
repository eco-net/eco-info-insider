<!--#include virtual="/shared/common.asp"-->
<!--#include virtual="/connections/ecoinfo.asp"-->

<%
DIM theJS
theJS=""

'1. Getting kommuner and creating arrays
DIM theRS,i,theid
SET theRS=Server.CreateObject("ADODB.Recordset")
theRS.activeconnection=MM_ecoinfo_STRING
theRS.source="SELECT k.nr,k.navn,a.nr as amt FROM eisys_kommuner k LEFT JOIN eisys_amter a ON k.amtnr=a.nr ORDER BY a.navn ASC"
theRS.Open()
i=0
theid=0
WHILE NOT theRS.EOF
	i=i+1
	IF theRS.Fields.Item("amt").value<>theid THEN
		IF i>1 THEN theJS = theJS & ")" & chr(13)
		theid=theRS.Fields.Item("amt").value
		theJS=theJS & "var array" & theRS.Fields.Item("amt").value & "=  new Array(""('<Vælg>','0',true,true)"""
	END IF
	theJS=theJS & ",""('" & theRS.Fields.Item("navn").value & "','" & theRS.Fields.Item("nr").value & "')"""
	theRS.movenext
WEND
IF i>0 THEN theJS=theJS & ")" & chr(13)
theRS.Close()


theJS=theJS & "function populateSel(theSel,selected) {" & chr(13) &_
	"	if(selected>0){" & chr(13) &_
	"		var selectedArray = eval(""array""+selected);" & chr(13) &_
	"		while (selectedArray.length < theSel.options.length) {" & chr(13) &_
	"			theSel.options[(theSel.options.length - 1)] = null;" & chr(13) &_
	"		}" & chr(13) &_
	"		for (var i=0; i < selectedArray.length; i++) {" & chr(13) &_
	"			theSel.options[i]=eval(""new Option"" + selectedArray[i]);" & chr(13) &_
	"		}" & chr(13) &_
	"	} else {" & chr(13) &_
	"		while (theSel.options.length>0) {" & chr(13) &_
	"			theSel.options[(theSel.options.length - 1)] = null;" & chr(13) &_
	"		}" & chr(13) &_
	"		theSel.options[0]=eval(""new Option"" + ""('<--','0',true,true)"");" & chr(13) &_
	"	}" & chr(13) &_
	"}"

WriteFile theJS,"/shared/selamtkomm.js"
%>