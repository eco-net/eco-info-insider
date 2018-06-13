<%
'**** Calculating sql-statement

DIM theFrom, theWhere, theOrder, searchDescr, eicount, eibackcount,einextcount,validtext
validtext=""
theFrom="eiorg_maindata m"
theWhere="0=0"
theOrder="ORDER BY m.zip"
searchDescr=""

eicount=CINT(request("count"))+0
IF eicount=0 THEN
	eibackcount=""
ELSE
	eibackcount=request("count")-1
END IF
einextcount=eicount+1

IF LEN(request("valid"))>0 THEN
	theWhere=theWhere & " AND m.validated=" & request("valid")
	
	IF request("valid")=0 THEN
		searchDescr=searchDescr & "<b>Ikke-godkendte</b> organisationer"
		validtext="IKKE GODKENDTE"
	ELSE
		searchDescr=searchDescr & "<b>Godkendte</b> organisationer"
	END IF
ELSEIF LEN(request("dgsid"))>0 then 'søgning på et eller flere ID-numre
	ids=replace(replace(replace(replace(request("dgsid"),chr(11),","),chr(13),","),chr(10),"")," ","")
	if right(ids,1)="," then ids=ids&"0"
	theWhere=theWhere & " AND m.fmpid in (" & ids & ")"
	searchDescr=searchDescr & "Organisationer med <b>bestemte Filemaker ID-numre</b>"
	validtext="MED BESTEMT Filemaker ID-nr"
ELSEIF LEN(request("netid"))>0 then 'søgning på et eller flere ID-numre
	ids=replace(replace(replace(replace(request("netid"),chr(11),","),chr(13),","),chr(10),"")," ","")
	if right(ids,1)="," then ids=ids&"0"
	theWhere=theWhere & " AND m.id in (" & ids & ")"
	searchDescr=searchDescr & "Organisationer med <b>bestemte Net ID-numre</b>"
	validtext="MED BESTEMT Net ID-nr"
ELSEIF LEN(request("stopped"))>0 then 'søgning på ophørte
	theWhere=theWhere & " AND m.stopped=1"
	searchDescr=searchDescr & "<strong>Ophørte</strong> organisationer"
	validtext="der er ophørte"
ELSE
	IF request("dgscat")>0 THEN
		theWhere=theWhere & " AND oc.catid=" & request("dgscat")
		theFROM= "(" & theFrom & " LEFT JOIN eiorg_r_cat oc ON m.id=oc.orgid)"
		searchDescr=searchDescr & " i kategorien <b>" & request("dgscatname") & "</b>"
	END IF
	IF request("dgssbj")>0 THEN
		theWhere=theWhere & " AND os.sbjid=" & request("dgssbj")
		theFROM= "(" & theFrom & " LEFT JOIN eiorg_r_sbj os ON m.id=os.orgid)"
		searchDescr=searchDescr & " og tilknyttet emnet <b>" & request("dgssbjname") & "</b>"
	END IF
	IF request("dgskomm")>0 THEN
		theWhere=theWhere & " AND p.komnr=" & request("dgskomm")
		theFROM= "(" & theFrom & " LEFT JOIN eisys_postnr p ON m.zip_dk=p.postnr)"
		searchDescr=searchDescr & " og hjemh&oslash;rende i kommunen <b>" & request("dgskommname") & "</b>"
	ELSEIF request("dgsamt")>0 THEN
		theWhere=theWhere & " AND p.amtnr=" & request("dgsamt")
		theFROM= "(" & theFrom & " LEFT JOIN eisys_postnr p ON m.zip_dk=p.postnr)"
		searchDescr=searchDescr & " og hjemh&oslash;rende i amtet <b>" & request("dgsamtname") & "</b>"
	END IF
	IF LEN(request("dgstext"))>0 THEN
		theWhere=theWhere & " AND (m.name LIKE '%" & request("dgstext") &_
			"%' OR m.fullname LIKE '%" & request("dgstext") &_
			"%' OR m.shortdescription LIKE '%" & request("dgstext") &_
			"%' OR m.keywords LIKE '%" & request("dgstext") &_
			"%' OR m.description LIKE '%" & request("dgstext") & "%')"
		searchDescr=searchDescr & " og hvor navnet eller beskrivelsen indeholder teksten <b>" & request("dgsText") & "</b>"
	END IF
	
		
END IF
IF searchDescr="" THEN
	searchDescr="Alle organisationer"
ELSE
	IF LEFT(searchDescr,2)=" o" THEN
		searchDescr="Organisationer " & RIGHT(searchDescr,LEN(searchDescr)-3) & "."
	ELSEIF LEFT(searchDescr,1)="<" THEN
		searchDescr=searchDescr & "."
	ELSE
		searchDescr="Organisationer " & searchDescr & "."
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
