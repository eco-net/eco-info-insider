<%
'**** Calculating sql-statement

DIM theFrom, theWhere, searchDescr, eicount, eibackcount,einextcount,theOrder,validtext
validtext=""
theFrom="eiudd_maindata m"
theWhere="0=0"
theOrder="ORDER BY m.title"
searchDescr=""

eicount=CINT(request("count"))+0
IF eicount=0 THEN
	eibackcount=""
ELSE
	eibackcount=request("count")-1
END IF
einextcount=eicount+1

IF LEN(request.cookies("eiorgid"))>0 THEN

	theWhere=theWhere & " AND ud.orgid=" & request.cookies("eiorgid")
	theFROM= "(" & theFrom & " LEFT JOIN eiudd_r_org ud ON m.id=ud.uddid)"
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & request.cookies("eiorgname") & "</b>"
ELSEIF LEN(request("orgid"))>0 THEN

	theWhere=theWhere & " AND ud.orgid=" & request("orgid")
	theFROM= "(" & theFrom & " LEFT JOIN eiudd_r_org ud ON m.id=ud.uddid)"
	searchDescr=searchDescr & " tilknyttet organisationen <b>" & request("orgname") & "</b>"
ELSEIF LEN(request("valid"))>0 THEN
	theWhere=theWhere & " AND m.validated=" & request("valid")
	IF request("valid")=0 THEN
		searchDescr=searchDescr & "<b>Ikke-godkendte</b> uddannelser"
		validtext="IKKE GODKENDTE"
	ELSE
		searchDescr=searchDescr & "<b>Godkendte</b> uddannelser"
	END IF
ELSE

	IF request("udcat")>0 THEN
		theWhere=theWhere & " AND cc.catid=" & request("udcat")
		theFROM= "(" & theFrom & " LEFT JOIN eiudd_r_cat cc ON m.id=cc.uddid)"
		searchDescr=searchDescr & " i kategorien <b>" & request("udcatname") & "</b>"
	END IF
	IF request("udsbj")>0 THEN
		theWhere=theWhere & " AND cs.sbjid=" & request("udsbj")
		theFROM= "(" & theFrom & " LEFT JOIN eiudd_r_sbj cs ON m.id=cs.uddid)"
		searchDescr=searchDescr & " og tilknyttet emnet <b>" & request("udsbjname") & "</b>"
	END IF
	IF LEN(request("udtext"))>0 THEN
		theWhere=theWhere & " AND (m.title LIKE '%" & request("udtext") &_
			"%' OR m.shortdescription LIKE '%" & request("udtext") &_
			"%' OR m.description LIKE '%" & request("udtext") & "%')"
		searchDescr=searchDescr & " og hvor titel eller beskrivelse indeholder teksten <b>" & request("udtext") & "</b>"
	END IF
END IF
IF searchDescr="" THEN
	searchDescr="Alle uddannelser"
ELSE
	IF LEFT(searchDescr,2)=" o" THEN
		searchDescr="Uddannelser " & RIGHT(searchDescr,LEN(searchDescr)-3) & "."
	ELSEIF LEFT(searchDescr,1)="<" THEN
		searchDescr=searchDescr & "."
	ELSE
		searchDescr="Uddannelser " & searchDescr & "."
	END IF	
END IF


%>
