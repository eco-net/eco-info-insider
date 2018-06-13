<%
'**** Calculating sql-statement

DIM theFrom,theWhere,theOrder,searchDescr,eicount,eibackcount,einextcount,header
theFrom="eilib_maindata m"
theWhere="0=0"
'CreateDate > " + Replace(rsPageData__off, "'", "''") + ""
searchDescr=""
theOrder=""
eicount=CINT(request("count"))+0
IF eicount=0 THEN
	eibackcount=""
ELSE
	eibackcount=request("count")-1
END IF
einextcount=eicount+1

IF LEN(request.cookies("eiorgid"))>0 THEN
	theWhere=theWhere & " AND lo.orgid=" & request.cookies("eiorgid")
	theFrom= "(" & theFrom & " LEFT JOIN eilib_r_org lo ON m.id=lo.libid)"
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & request.cookies("eiorgname") & "</b>"
ELSEIF LEN(request("orgid"))>0 THEN
	theWhere=theWhere & " AND lo.orgid=" & request("orgid")
	theFrom= "(" & theFrom & " LEFT JOIN eilib_r_org lo ON m.id=lo.libid)"
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & request("orgname") & "</b>"
ELSEIF LEN(request("valid"))>0 THEN
	theWhere=theWhere & " AND Approved=" & request("valid")
	IF request("valid")=0 THEN
		searchDescr=searchDescr & "<b>Ikke-godkendte</b> nyheder"
		header="Nyheder der <font color='red'>ikke er godkendte</font>"
	ELSE
		searchDescr=searchDescr & "<b>Godkendte</b> nyheder"
		
	END IF
ELSE
	header="Alle Nyheder"
	IF request("dgbcat")>0 THEN
		theWhere=theWhere & " AND lc.catid=" & request("dgbcat")
		theFrom= "(" & theFrom & " LEFT JOIN eilib_r_cat lc ON m.id=lc.libid)"
		searchDescr=searchDescr & " i kategorien <b>" & request("dgbcatname") & "</b>"
	END IF
	IF request("dgbsbj")>0 THEN
		theWhere=theWhere & " AND ls.sbjid=" & request("dgbsbj")
		theFrom= "(" & theFrom & " LEFT JOIN eilib_r_sbj ls ON m.id=ls.libid)"
		searchDescr=searchDescr & " og tilknyttet emnet <b>" & request("dgbsbjname") & "</b>"
	END IF
	IF LEN(request("dgbtext"))>0 THEN
		theWhere=theWhere & " AND (m.title LIKE '%" & request("dgbtext") &_
			"%' OR m.subtitle LIKE '%" & request("dgbtext") &_
			"%' OR m.shortdescription LIKE '%" & request("dgbtext") &_
			"%' OR m.description LIKE '%" & request("dgbtext") & "%')"
		searchDescr=searchDescr & " og hvor navnet eller beskrivelsen indeholder teksten <b>" & request("dgbText") & "</b>"
	END IF
END IF
IF searchDescr="" THEN
	searchDescr="Alle Nyheder"
ELSE
	IF LEFT(searchDescr,2)=" o" THEN
		searchDescr="Nyheder " & RIGHT(searchDescr,LEN(searchDescr)-3) & "."
	ELSEIF LEFT(searchDescr,1)="<" THEN
		searchDescr=searchDescr & "."
	ELSE	
		searchDescr="Nyheder " & searchDescr & "."
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
