<%
'**** Calculating sql-statement

DIM theFrom,theWhere,theOrder,searchDescr,eicount,eibackcount,einextcount,validtext
validtext=""
theFrom="eiart_maindata m"
theWhere="0=0"
searchDescr=""
theOrder="ORDER BY m.title"
eicount=CINT(request("count"))+0
IF eicount=0 THEN
	eibackcount=""
ELSE
	eibackcount=request("count")-1
END IF
einextcount=eicount+1

IF LEN(request.cookies("eiorgid"))>0 THEN
	theWhere=theWhere & " AND lo.orgid=" & request.cookies("eiorgid")
	theFrom= "(" & theFrom & " LEFT JOIN eiart_r_org lo ON m.id=lo.artid)"
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & request.cookies("eiorgname") & "</b>"
ELSEIF LEN(request("orgid"))>0 THEN
	theWhere=theWhere & " AND lo.orgid=" & request("orgid")
	theFrom= "(" & theFrom & " LEFT JOIN eiart_r_org lo ON m.id=lo.artid)"
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & request("orgname") & "</b>"
ELSEIF LEN(request("valid"))>0 THEN
	theWhere=theWhere & " AND m.validated=" & request("valid")
	IF request("valid")=0 THEN
		searchDescr=searchDescr & "<b>Ikke-godkendte</b> artikler"
		validtext="IKKE GODKENDTE"
	ELSE
		searchDescr=searchDescr & "<b>Godkendte</b> artikler"
	END IF
ELSE
	IF request("artcat")>0 THEN
		theWhere=theWhere & " AND lc.catid=" & request("artcat")
		theFrom= "(" & theFrom & " LEFT JOIN eiart_r_cat lc ON m.id=lc.artid)"
		searchDescr=searchDescr & " i kategorien <b>" & request("artcatname") & "</b>"
	END IF
	IF request("artsbj")>0 THEN
		theWhere=theWhere & " AND ls.sbjid=" & request("artsbj")
		theFrom= "(" & theFrom & " LEFT JOIN eiart_r_sbj ls ON m.id=ls.artid)"
		searchDescr=searchDescr & " og tilknyttet emnet <b>" & request("artsbjname") & "</b>"
	END IF
	IF LEN(request("arttext"))>0 THEN
		theWhere=theWhere & " AND (m.title LIKE '%" & request("arttext") &_
			"%' OR m.subtitle LIKE '%" & request("arttext") &_
			"%' OR m.shortdescr LIKE '%" & request("arttext") &_
			"%' OR m.descr LIKE '%" & request("arttext") & "%')"
		searchDescr=searchDescr & " og hvor navnet eller beskrivelsen indeholder teksten <b>" & request("artText") & "</b>"
	END IF
END IF
IF searchDescr="" THEN
	searchDescr="Alle artikler"
ELSE
	IF LEFT(searchDescr,2)=" o" THEN
		searchDescr="Artikler " & RIGHT(searchDescr,LEN(searchDescr)-3) & "."
	ELSEIF LEFT(searchDescr,1)="<" THEN
		searchDescr=searchDescr & "."
	ELSE	
		searchDescr="Artikler " & searchDescr & "."
	END IF	
END IF

FUNCTION FormatDateDB(theDate)
	theDate=CDate(theDate)
	FormatDateDB="{ts '" & datepart("yyyy",theDate) & "-" & right("0" & CStr(datepart("m",theDate)),2) & "-" &_
		right("0" & CSTR(datepart("d",theDate)),2) & " 00:00:00'}"
END FUNCTION

'response.write theWhere & "<br>"
'response.write theFROM & "<br>"
'response.end
%>
