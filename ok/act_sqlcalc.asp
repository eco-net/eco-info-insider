<%
'**** Calculating sql-statement

DIM theFrom, theWhere, searchDescr, eicount, eibackcount,einextcount,theOrder,validtext
validtext=""
theFrom="(eical_maindata m LEFT JOIN eisys_postnr p ON m.postnr=p.postnr)"
theWhere="0=0"
IF LEN(request.cookies("eiuserid"))>0 THEN
	theOrder="ORDER BY m.startdate desc"
ELSE
	theOrder="ORDER BY m.startdate DESC"
END IF
searchDescr=""

eicount=CINT(request("count"))+0
IF eicount=0 THEN
	eibackcount=""
ELSE
	eibackcount=request("count")-1
END IF
einextcount=eicount+1

IF LEN(request.cookies("eiorgid"))>0 THEN
	theWhere=theWhere & " AND co.orgid=" & request.cookies("eiorgid")
	theFROM= "(" & theFrom & " LEFT JOIN eical_r_org co ON m.id=co.calid)"
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & request.cookies("eiorgname") & "</b>"
ELSEIF LEN(request("orgid"))>0 THEN
	theWhere=theWhere & " AND co.orgid=" & request("orgid")
	theFROM= "(" & theFrom & " LEFT JOIN eical_r_org co ON m.id=co.calid)"
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & request("orgname") & "</b>"
ELSEIF LEN(request("valid"))>0 THEN
	theWhere=theWhere & " AND m.validated=" & request("valid")
	IF request("valid")=0 THEN
		searchDescr=searchDescr & "<b>Ikke-godkendte</b> arrangementer"
		validtext="IKKE GODKENDTE"
	ELSE
		searchDescr=searchDescr & "<b>Godkendte</b> arrangementer"
	END IF
ELSE
	IF LEN(request("oktime"))>1 THEN
		DIM theDate
		IF CSTR(request("oktime"))="10" THEN
		'Kommende
			theWhere=theWhere & " AND (m.startdate>=" & FormatDateDB(date) &_
				" OR m.enddate>=" & FormatDateDB(date) & ")"
			searchDescr=searchDescr & "<b>Kommende</b> arrangementer"
		ELSEIF CSTR(request("oktime"))="20" THEN
		'Afholdte
			theWhere=theWhere & " AND m.startdate<=" & FormatDateDB(date) &_
				" AND m.enddate<=" & FormatDateDB(date)
			searchDescr=searchDescr & "<b>Afholdte</b> arrangementer"
		ELSEIF CSTR(request("oktime"))="30" THEN
		'denne måned
			theDate=DateSerial(DatePart("yyyy",date),DatePart("m",date)+1,1)
			theDate=DateAdd("d",-1,theDate)
			theDate=FormatDateDB(theDate)
			theWhere=theWhere & " AND (m.startdate BETWEEN " & FormatDateDB(Date) &_
				" AND " & theDate & " OR " &_
				"m.enddate BETWEEN " & FormatDateDB(Date) &_
				" AND " & theDate & ")"
			searchDescr=searchDescr & " i <b>resten af denne måned</b>"
		ELSEIF CSTR(request("oktime"))="40" THEN
		'næste måned
			DIM startDate
			startDate=DateSerial(DatePart("yyyy",date),DatePart("m",date)+1,1)
			theDate=DateAdd("m",1,startDate)
			startDate=FormatDateDB(startDate)
			theDate=DateAdd("d",-1,theDate)
			theDate=FormatDateDB(theDate)
			theWhere=theWhere & " AND (m.startdate BETWEEN " & startDate &_
				" AND " & theDate & " OR " &_
				"m.enddate BETWEEN " & startDate &_
				" AND " & theDate & ")"
			searchDescr=searchDescr & " i <b>næste måned</b>"
		ELSE
		'en måned
			theDate=DateAdd("m",1,CDate(request("oktime")))
			theDate=DateAdd("d",-1,theDate)
			theDate=FormatDateDB(theDate)
			theWhere=theWhere & " AND (m.startdate BETWEEN " & FormatDateDB(request("oktime")) &_
				" AND " & theDate & " OR " &_
				"m.enddate BETWEEN " & FormatDateDB(request("oktime")) &_
				" AND " & theDate & ")"
			searchDescr=searchDescr & " i perioden <b>" & request("oktimename") & "</b>"
		END IF
	ELSE
	'Kommende
		theWhere=theWhere & " AND (m.startdate>=" & FormatDateDB(date) &_
			" OR m.enddate>=" & FormatDateDB(date) & ")"
		searchDescr=searchDescr & "<b>Kommende</b> arrangementer"
	END IF
	IF request("okcat")>0 THEN
		theWhere=theWhere & " AND cc.catid=" & request("okcat")
		theFROM= "(" & theFrom & " LEFT JOIN eical_r_cat cc ON m.id=cc.calid)"
		searchDescr=searchDescr & " i kategorien <b>" & request("okcatname") & "</b>"
	END IF
	IF request("oksbj")>0 THEN
		theWhere=theWhere & " AND cs.sbjid=" & request("oksbj")
		theFROM= "(" & theFrom & " LEFT JOIN eical_r_sbj cs ON m.id=cs.calid)"
		searchDescr=searchDescr & " og tilknyttet emnet <b>" & request("oksbjname") & "</b>"
	END IF
	IF request("okkomm")>0 THEN
		theWhere=theWhere & " AND p.komnr=" & request("okkomm")
		searchDescr=searchDescr & " og afholdes i kommunen <b>" & request("okkommname") & "</b>"
	ELSEIF request("okamt")>0 THEN
		theWhere=theWhere & " AND p.amtnr=" & request("okamt")
		searchDescr=searchDescr & " og afholdes i amtet <b>" & request("okamtname") & "</b>"
	END IF
	IF LEN(request("oktext"))>0 THEN
		theWhere=theWhere & " AND (m.title LIKE '%" & request("oktext") &_
			"%' OR m.shortdescription LIKE '%" & request("oktext") &_
			"%' OR m.description LIKE '%" & request("oktext") & "%')"
		searchDescr=searchDescr & " og hvor titel eller beskrivelse indeholder teksten <b>" & request("okText") & "</b>"
	END IF
END IF
IF searchDescr="" THEN
	searchDescr="Alle arrangementer"
ELSE
	IF LEFT(searchDescr,2)=" o" THEN
		searchDescr="Arrangementer " & RIGHT(searchDescr,LEN(searchDescr)-3) & "."
	ELSEIF LEFT(searchDescr,1)="<" THEN
		searchDescr=searchDescr & "."
	ELSE
		searchDescr="Arrangementer " & searchDescr & "."
	END IF	
END IF
'response.write theWhere & "<br>"
'response.write theFROM & "<br>"
'response.end

FUNCTION FormatDateDB(theDate)
	theDate=CDate(theDate)
	FormatDateDB="{ts '" & datepart("yyyy",theDate) & "-" & right("0" & CStr(datepart("m",theDate)),2) & "-" &_
		right("0" & CSTR(datepart("d",theDate)),2) & " 00:00:00'}"
END FUNCTION
%>
