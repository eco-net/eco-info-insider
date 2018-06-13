<!--include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<!--#include virtual="/shared/insiderusers.asp" -->
<%

Dim theID
SUB DoSave(act)

	'**** Creating SQL-statement
	IF act=1 THEN 'Insert
		DoCreate
		
		regUser request("username"),theID,act,"dgs",theID,"stamdata"
	ELSEIF act=2 THEN 'Update
		DoUpdate
		regUser Session("username"),request("id"),act,"dgs",request("id"),"stamdata"
	ELSEIF act=3 THEN 'Delete
		DoDelete
	END IF

	'**** Showing confirmation to user
	IF INstr(request("params"),"newuser")>0 THEN
		response.redirect("welcome.asp")	
	ELSE
		response.redirect("confirmation.asp?done=" & act & request("params"))
	END IF
END SUB

SUB DoCreate

	status=2
	creator=2
		

	valid=0
	if (request.cookies("eiorgname") = "&Oslash;ko-net medarbejder") then 
		status=1
		creator=1
		valid=1
	end if

	theCity=GetCity(request("zip_dk"))
	fullname=request("firstname") & " " & request("lastname")
	'1. Inserting org.
	theSQL="INSERT INTO eiorg_maindata(" &_
	"name,adrco,street1,street2,zip,zip_dk,city,phone1,phone2,fax,emailaddress1,emailaddress2,www,www2,shortdescription,description,firstname,lastname,title,validated,keywords,status,creator,fullname,maker,phone3" &_
	") VALUES (" &_
	"'" & replace(request("name"),"'","''") & "'" & "," &_
	"'" & replace(request("adrco"),"'","''") & "'" & "," &_
	"'" & replace(request("street1"),"'","''") & "'" & "," &_
	"'" & replace(request("street2"),"'","''") & "'" & "," &_
	"'" & request("zip") & "'" & "," &_
	request("zip_dk") & "," &_
	"'" & theCity & "'" & "," &_
	"'" & request("phone1") & "'" & "," &_
	"'" & request("phone2") & "'" & "," &_
	"'" & request("fax") & "'" & "," &_
	"'" & request("emailaddress1") & "'" & "," &_
	"'" & request("emailaddress2") & "'" & "," &_
	"'" & request("www") & "'" & "," &_
	"'" & request("www2") & "'" & "," &_
	"'" & replace(request("shortdescription"),"'","''") & "'" & "," &_
	"'" & replace(request("description"),"'","''") & "'" & "," &_
	"'" & replace(request("firstname"),"'","''") & "'" & "," &_
	"'" & replace(request("lastname"),"'","''") & "'" & "," &_
	"'" & replace(request("title"),"'","''") & "'" & "," &_
	valid & "," &_
	"'" & replace(request("keywords"),"'","''") & "'" & ","  &_
	status  & "," &_
	creator & "," &_
	"'" & fullname & "'" & "," &_
	"'" & Session("username") & "'" & "," &_
	"'" & request("phone3") & "'" &_
	")"
	
	theID=InsertRec(theSQL,"eiorg_maindata","id","")
	
	'2. Inserting Categories.
	theSQL=""
	FOR each Item in request("selcat")
		theSQL=theSQL & "INSERT INTO eiorg_r_cat (orgid,catid) VALUES (" & theID & "," & item & ");" & chr(13)
	NEXT
	execCommand theSQL
	
	'3. Inserting Subjects
	theSQL=""
	FOR each Item in request("selsbj")
		theSQL=theSQL & "INSERT INTO eiorg_r_sbj (orgid,sbjid) VALUES (" & theID & "," & item & ");" & chr(13)
	NEXT
	execCommand theSQL

	'4. Inserting User
	userpw=GetUniquepw(request("username"),request("password"))
	theSQL="INSERT INTO eisys_insiderusers(" &_
	"username,password,orgid,email,userlevel" &_
	") VALUES (" &_
	"'" & request("username") & "'" & "," &_
	"'" & userpw & "'" & "," &_
	theID & "," &_
	"'" & request("email") & "'" & "," &_
	"0"  &_
	")"
	
	execCommand theSQL
	
	'5. Sending mail to user
	theBody=ReadFile("/admin/ei/snippets/velkomstmail.txt")
	theBody=replace(theBody,"#username#",request("username"))
	theBody=replace(theBody,"#password#",userpw)
	theBody=replace(theBody,"#orgid#",theID)
	theBody=replace(theBody,"#title#",request("name"))
	'SendMail request("email"),"eco-net@eco-net.dk","Tak for din tilmelding til ko-info",theBody
	
	SendCDOMail request("email"),"eco-net@eco-net.dk","Tak for din tilmelding til ko-info",theBody
	'6 Sending mail to Econet
	theBody="Kopier nedenstende til oprettelse i Filemaker:" & VbCrLf
	theBody=theBody & "Navn:" & request("name") & ";Fornavn:" & request("firstname") & ";Efternavn:" & request("lastname") & ";Mail:" &_
		request("emailaddress1") & ";Gade:" & request("street1") & ";Postnr:" & request("zip") & ";By:" &_
		theCity & ";Gade2:" & request("street2") & ";Tlf:" & request("phone1") & ";Tlf2:" & request("phone2") &_
		";Tlf3:" & request("phone3") & ";Fax:" & request("fax") & ";WWW:" & request("www") & ";WWW2:" & request("www2") & ";Title:" & request("title") & ";SQLID:" & theID &_
		";CO:" & request("adrco") & ";"

	theBody=theBody & VbCrLf & VbCrLf & "Afsendt den " & FormatDateTime(Date(),vbLongDate) & " kl. " & FormatDateTime(Now(),vbShortTime)
	'SendMail "eco-net@eco-net.dk","eco-net@eco-net.dk","Tilmelding p De Grnne Sider",theBody
	SendCDOMail "eco-net@eco-net.dk","eco-net@eco-net.dk","Tilmelding p De Grnne Sider",theBody
	DoCount
END SUB

SUB DoUpdate
	status=4
	valid=0
	if (request.cookies("eiorgname") = "&Oslash;ko-net medarbejder") then 
		status=3
		valid=1
	end if
	'1. Updating org
	'IF LEN(request("validated"))>0 THEN
	'	validated="1"
	'ELSE
	'	validated="0"
	'END IF
	IF LEN(request("stopped"))>0 THEN
		stopped="1"
	ELSE
		stopped="0"
	END IF
	theCity=GetCity(request("zip_dk"))
	fullname=request("firstname") & " " & request("lastname")
	theSQL="UPDATE eiorg_maindata SET " &_
	"name='" & replace(request("name"),"'","''") & "'" & "," &_
	"street1='" & replace(request("street1"),"'","''") & "'" & "," &_
	"street2='" & replace(request("street2"),"'","''") & "'" & "," &_
	"adrco='" & replace(request("adrco"),"'","''") & "'" & "," &_
	"zip='" & request("zip") & "'" & "," &_
	"zip_dk=" & request("zip_dk") & "," &_
	"city='" & theCity & "'" & "," &_
	"phone1='" & request("phone1") & "'" & "," &_
	"phone2='" & request("phone2") & "'" & "," &_
	"phone3='" & request("phone3") & "'" & "," &_
	"fax='" & request("fax") & "'" & "," &_
	"emailaddress1='" & request("emailaddress1") & "'" & "," &_
	"emailaddress2='" & request("emailaddress2") & "'" & "," &_
	"www='" & request("www") & "'" & "," &_
	"www2='" & request("www2") & "'" & "," &_
	"shortdescription='" & replace(request("shortdescription"),"'","''") & "'" & "," &_
	"description='" & replace(request("description"),"'","''") & "'" & "," &_
	"firstname='" & replace(request("firstname"),"'","''") & "'" & "," &_
	"lastname='" & replace(request("lastname"),"'","''") & "'" & "," &_
	"fullname='" & fullname & "'" & "," &_
	"title='" & replace(request("title"),"'","''") & "'" & "," &_
	"validated=" & valid & "," &_
	"modidate=" & FormatDateDB(date) & "," &_
	"keywords='" & replace(request("keywords"),"'","''") & "'" & "," &_
	"status=" & status & "," &_
	"editor='" & Session("username") & "'" & "," &_
	"stopped=" & stopped &_
	" WHERE id=" & request("id")
	execCommand theSQL

	'2. Updating categories
	IF LEN(request("selCat"))>0 THEN
		theSQL="DELETE FROM eiorg_r_cat WHERE orgid=" & request("id")
		execCommand theSQL
		theSQL=""
		FOR each Item in request("selcat")
			theSQL=theSQL & "INSERT INTO eiorg_r_cat (orgid,catid) VALUES (" & request("id") & "," & item & ");" & chr(13)
		NEXT
		execCommand theSQL
	END IF

	'3. Updating subjects
	IF LEN(request("selSbj"))>0 THEN
		theSQL="DELETE FROM eiorg_r_sbj WHERE orgid=" & request("id")
		execCommand theSQL
		theSQL=""
		FOR each Item in request("selsbj")
			theSQL=theSQL & "INSERT INTO eiorg_r_sbj (orgid,sbjid) VALUES (" & request("id") & "," & item & ");" & chr(13)
		NEXT
		execCommand theSQL
	END IF
	DoCount
	if stopped = "1" then
		theSQL="UPDATE eisys_insiderusers SET " &_
		"password='slettet'" & "," &_
		"username='slettet'" & "," &_
		"email=''" &_
		" WHERE orgid=" & request("id") 
		
		execCommand theSQL
	end if

	'4. Inserting User
	IF len(request("username")) then
		userpw=GetUniquepw(request("username"),request("password"))
		theSQL="INSERT INTO eisys_insiderusers(" &_
		"username,password,orgid,email,userlevel" &_
		") VALUES (" &_
		"'" & request("username") & "'" & "," &_
		"'" & userpw & "'" & "," &_
		request("id") & "," &_
		"'" & request("emailaddress1") & "'" & "," &_
		"0"  &_
		")"
		execCommand theSQL
	else
		' opdater insiderusers s mailadresse er aktuel
		theSQL="UPDATE eisys_insiderusers SET email='" & request("emailaddress1") & "' WHERE orgid=" & request("id") 
				execCommand(theSQL)
	end if

END SUB

SUB DoDelete
	theid=request("id")
	theSQL="Delete from eiorg_r_sbj WHERE orgid=" & theid & ";" & chr(13)
	theSQL=theSQL & "Delete from eiorg_r_cat WHERE orgid=" & theid & ";" & chr(13)
	theSQL=theSQL & "Delete from eisys_insiderusers WHERE orgid=" & theid & ";" & chr(13)
	theSQL=theSQL & "Delete from eiorg_r_coll WHERE orgid=" & theid & ";" & chr(13)
	theSQL=theSQL & "Delete from eiorg_maindata WHERE id=" & theid & ";" & chr(13)
	execCommand theSQL

	set rs=server.createobject("adodb.recordset")
	rs.activeconnection=MM_ecoinfo_STRING
	rs.source="Select * FROM eilib_r_org WHERE orgid=" & theid
	rs.open
	WHILE NOT rs.EOF
		libid=rs("libid")
		theSQL="Delete FROM eilib_r_sbj WHERE libid=" & libid & ";" & chr(13)
		theSQL=theSQL & "Delete FROM eilib_r_cat WHERE libid=" & libid & ";" & chr(13)
		theSQL=theSQL & "Delete FROM eilib_r_coll WHERE libid=" & libid & ";" & chr(13)
		theSQL=theSQL & "Delete FROM eilib_maindata WHERE id=" & libid & ";" & chr(13)
		execCommand theSQL
		rs.movenext
	WEND
	theSQL="Delete from eilib_r_org WHERE orgid=" & theid
	rs.close()

	rs.source="Select * FROM eilib_r_org WHERE orgid=" & theid
	rs.open
	WHILE NOT rs.EOF
		libid=rs("calid")
		theSQL="Delete FROM eical_r_sbj WHERE calid=" & libid & ";" & chr(13)
		theSQL=theSQL & "Delete FROM eical_r_cat WHERE calid=" & libid & ";" & chr(13)
		theSQL=theSQL & "Delete FROM eical_r_coll WHERE calid=" & libid & ";" & chr(13)
		theSQL=theSQL & "Delete FROM eical_maindata WHERE id=" & libid & ";" & chr(13)
		execCommand theSQL
		rs.movenext
	WEND
	theSQL="Delete from eical_r_org WHERE orgid=" & theid
	rs.close()
	set rs=nothing
	
	'UpdateIncludes
END SUB

SUB UpdateIncludes
	MakeMenuFile "SELECT id,headline FROM eiban_maindata WHERE siteid=1 ORDER BY 2","&lt;V&aelig;lg&gt;","id","headline","/log/insider/ei/menufiles/sel_ban_dyn.txt","cursel"
	MakeMenuFile "SELECT id,headline FROM eiban_maindata WHERE siteid=1 ORDER BY 2","&lt;V&aelig;lg&gt;","id","headline","/log/insider/ei/menufiles/sel_ban.txt",""
END SUB

SUB DoCount
	DIM rsRecCount, organisations, theText, filename
	set rsRecCount = Server.CreateObject("ADODB.Recordset")
	rsRecCount.ActiveConnection = MM_ecoinfo_STRING
	'rsRecCount.Source = "SELECT Count(*) as theCount FROM eiorg_maindata m WHERE m.validated=1"
	rsRecCount.Source = "SELECT Count(*) as theCount FROM eiorg_maindata m "
	rsRecCount.CursorType = 0
	rsRecCount.CursorLocation = 2
	rsRecCount.LockType = 3
	rsRecCount.Open()
	organisations = rsRecCount.Fields.Item("theCount").value
	rsRecCount.Close()
	theText = Chr(60) & Chr(37) & " organisations = " & organisations & Chr(37) & Chr(62)
	theFilename = "/log/ei/dgs/org_count.asp"
	
'	WriteFile theText,theFilename
	
END SUB

FUNCTION GetCity(zip)
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_ecoinfo_STRING
	rs.Source = "SELECT city FROM eisys_postnr WHERE postnr=" & zip
	rs.open()
	IF NOT rs.EOF THEN
		GetCity=replace(rs("city"),"'","''")
	ELSE
		GetCity="Ukendt"
	END IF
END FUNCTION

SUB DoInsert
	theID=InsertRec(request("SQLcode"),"eiorg_maindata","id","")

	'4. Inserting User
	userpw=GetUniquepw(request("username"),request("password"))
	theSQL="INSERT INTO eisys_insiderusers(" &_
	"username,password,orgid,email,userlevel" &_
	") VALUES (" &_
	"'" & request("username") & "'" & "," &_
	"'" & userpw & "'" & "," &_
	theID & "," &_
	"'" & request("email") & "'" & "," &_
	"0"  &_
	")"
	execCommand theSQL

	'5. Sending mail to user
	set rs=GetRecordSet("Select name from eiorg_maindata where id="&theID)
	theBody=ReadFile("/admin/ei/snippets/velkomstmail_byus.txt")
	theBody=replace(theBody,"#username#",request("username"))
	theBody=replace(theBody,"#password#",userpw)
	theBody=replace(theBody,"#orgid#",theID)
	theBody=replace(theBody,"#title#",rs("name"))
	set rs=nothing
	'SendMail request("email"),"eco-net@eco-net.dk","Du/I er blevet oprettet på Øko-info",theBody
	SendCDOMail request("email"),"eco-net@eco-net.dk","Du/I er blevet oprettet på Øko-info",theBody
	response.redirect "newfromfmp.asp?recid=" & theid
END SUB
%>