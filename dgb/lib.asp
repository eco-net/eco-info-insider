<!--include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<!--#include virtual="/shared/insiderusers.asp" -->
<%
Dim theorgID,theID
theorgID=request("orgid")
SUB DoSave(act)
	'**** Creating SQL-statement
	IF act=1 THEN 'Insert
		DoCreate
		'regUser Session("username"),theorgID,act,"dgb",theID,request("title")
	ELSEIF act=2 THEN 'Update
		DoUpdate
		regUser Session("username"),theorgID,act,"dgb",theID,request("title")
	ELSEIF act=3 THEN 'Delete
		DoDelete
	END IF

	'**** Showing confirmation to user
	response.redirect "confirmation.asp?done=" & act & request("params")
	response.end
END SUB

SUB DoCreate
	if request("img_id")<>"" then
		img_id=CInt(request("img_id"))
	else
		img_id=0
	end if
	if request("img_id2")<>"" then
		img_id2=CInt(request("img_id2"))
	else
		img_id2=0
	end if
	status=2
	if (request.cookies("eiorgname") = "&Oslash;ko-net medarbejder") then 
		status=1
	end if
	'1. Inserting org.
	theSQL="INSERT INTO eilib_maindata(" &_
	"title,subtitle,shortdescription,description,language,author,year,isbn,editor,editoremail,editorwww,webaddress,validated,keywords,status,maker,img_id,img_id2,islocal,imagefilename1,imagefilename2" &_
	") VALUES (" &_
	"'" & replace(request("title"),"'","''") & "'" & "," &_
	"'" & replace(request("subtitle"),"'","''") & "'" & "," &_
	"'" & replace(request("shortdescription"),"'","''") & "'" & "," &_
	"'" & replace(request("description"),"'","''") & "'" & "," &_
	"'" & request("language") & "'" & "," &_
	"'" & replace(request("author"),"'","''") & "'" & "," &_
	"'" & request("year") & "'" & "," &_
	"'" & request("isbn") & "'" & "," &_
	"'" & replace(request("editor"),"'","''") & "'" & "," &_
	"'" & request("editoremail") & "'" & "," &_
	"'" & request("editorwww") & "'" & "," &_
	"'" & request("webaddress") & "'" & "," &_
	"0" & "," &_
	"'" & replace(request("keywords"),"'","''") & "'"  & ","  &_
	status & "," &_
	"'" & Session("username") & "'"  & ","  &_
	img_id & "," &_
	img_id2 &","&_
	CINT("0"&request("islocal")) & ",'" &_
	request("imagefilename1") & "','" &_
	request("imagefilename2") & "')"
	theID=InsertRec(theSQL,"eilib_maindata","id","")
	
	'2. Inserting Categories.
	theSQL="INSERT INTO eilib_r_cat (libid,catid) VALUES (" & theID & "," & request("selCat") & ")"
	execCommand theSQL
	
	'3. Inserting Subjects
	theSQL=""
	FOR each Item in request("selsbj")
		theSQL=theSQL & "INSERT INTO eilib_r_sbj (libid,sbjid) VALUES (" & theID & "," & item & ");" & chr(13)
	NEXT
	execCommand theSQL

	'4. Linking to Org
	theSQL=""
	theorgs=split(request("orgid"),",")
	theorgID=theorgs(0)'for regUser()
	FOR i=0 to ubound(theorgs)
		theSQL=theSQL & "INSERT INTO eilib_r_org (orgid,libid) VALUES (" & theorgs(i) & "," & theID & ");" & chr(13)
	NEXT
	execCommand theSQL

	'5. Updating include-files
	'UpdateIncludes request("orgid")
	'response.end

	'6. Sender mail til org
	if LEN(request.cookies("eiuserid"))>0 then
		theBody=ReadFile("/admin/ei/snippets/dgbmail.txt")
		theBody=replace(theBody,"#titel#",request("title"))
		theBody=replace(theBody,"#id#",theID)
		for i=0 to ubound(theorgs)
			set rs=GetRecordSet("Select o.id as orgid, o.emailaddress1 as m, o.name, u.username,u.password from eiorg_maindata o inner join eisys_insiderusers u on o.id=u.orgid where o.id="&theorgs(i))
			tBody=replace(theBody,"#org#",rs("name"))
			tBody=replace(tBody,"#username#",rs("username"))
			tBody=replace(tBody,"#password#",rs("password"))
			tBody=replace(tBody,"#orgid#",rs("orgid"))
			'SendMail rs("m"),"eco-net@eco-net.dk","Vi har lagt Jeres publikation ind",tBody
				SendCDOMail rs("m"),"eco-net@eco-net.dk","Vi har lagt Jeres publikation ind",tBody

		next
		set rs=nothing
	end if
END SUB

SUB DoUpdate
	theID=request("id")
	if request("img_id")<>"" then
		img_id=CInt(request("img_id"))
	else
		img_id=0
	end if
	if request("img_id2")<>"" then
		img_id2=CInt(request("img_id2"))
	else
		img_id2=0
	end if
	status=4
	if (request.cookies("eiorgname") = "&Oslash;ko-net medarbejder") then 
		status=3
	end if

	'1. Updating org
	IF LEN(request("validated"))>0 THEN
		validated="1"
	ELSE
		validated="0"
	END IF
	
	theSQL="UPDATE eilib_maindata SET " &_
	"title='" & replace(request("title"),"'","''") & "'" & "," &_
	"subtitle='" & replace(request("subtitle"),"'","''") & "'" & "," &_
	"shortdescription='" & replace(request("shortdescription"),"'","''") & "'" & "," &_
	"description='" & replace(request("description"),"'","''") & "'" & "," &_
	"language='" & request("language") & "'" & "," &_
	"author='" & replace(request("author"),"'","''") & "'" & "," &_
	"year='" & request("year") & "'" & "," &_
	"isbn='" & request("isbn") & "'" & "," &_
	"editor='" & replace(request("editor"),"'","''") & "'" & "," &_
	"editoremail='" & request("editoremail") & "'" & "," &_
	"editorwww='" & request("editorwww") & "'" & "," &_
	"webaddress='" & request("webaddress") & "'" & "," &_
	"validated=" & validated & "," &_
	"modidate=" & FormatDateDB(date) & "," &_
	"keywords='" & replace(request("keywords"),"'","''") & "'" & "," &_
	"status=" & status & "," &_
	"edit='" & Session("username") & "'" & "," &_
	"img_id=" & img_id & "," &_
	"img_id2=" & img_id2 &","&_
	"islocal="&	CINT("0"&request("islocal")) & "," &_
	"imagefilename1='" & request("imagefilename1") & "'," &_
	"imagefilename2='" & request("imagefilename2") & "'" &_
	" WHERE id=" & request("id")
	
	execCommand theSQL

	'2. Updating categories
	IF LEN(request("selCat"))>0 THEN
		theSQL="UPDATE eilib_r_cat SET catid=" &  request("selCat") & " WHERE libid=" & request("id")
		execCommand theSQL
	END IF

	'3. Updating subjects
	IF LEN(request("selSbj"))>0 THEN
		theSQL="DELETE FROM eilib_r_sbj WHERE libid=" & request("id")
		execCommand theSQL
		theSQL=""
		FOR each Item in request("selsbj")
			theSQL=theSQL & "INSERT INTO eilib_r_sbj (libid,sbjid) VALUES (" & request("id") & "," & item & ");" & chr(13)
		NEXT
		execCommand theSQL
	END IF

	'4. Updating organizations
	IF len(request("theorgid"))>0 THEN
		theSQL = "DELETE FROM eilib_r_org WHERE libid=" & request("id")
		execCommand theSQL
		theSQL = ""
		theorgs=split(request("theorgid"),",")
		theOrgID = theorgs(0)
		FOR i=0 to ubound(theorgs)
			if LEN(theorgs(i))>1 then theSQL=theSQL & "INSERT INTO eilib_r_org (libid,orgid) VALUES (" & request("id") & "," & theorgs(i) & ");" & chr(13)
		NEXT
		execCommand theSQL
	END IF

	'4. Updating include-files
	'UpdateIncludes request("theorgid")

END SUB

SUB DoDelete
	theid=request("id")
	theSQL="Delete FROM eilib_r_sbj WHERE libid=" & theid
	execCommand theSQL
	theSQL="Delete FROM eilib_r_cat WHERE libid=" & theid
	execCommand theSQL
	theSQL="Delete FROM eilib_r_coll WHERE libid=" & theid
	execCommand theSQL
	theSQL="Delete FROM eilib_r_org WHERE libid=" & theid
	execCommand theSQL
	theSQL="Delete FROM eilib_maindata WHERE id=" & theid
	execCommand theSQL

	'4. Updating include-files
	'IF LEN(request("orgid"))>0 THEN UpdateIncludes request("orgid")
END SUB

SUB UpdateIncludes (orgid)
	starttext=""
	modelText="document.write('<font style=""font-size:larger""><b><a href=""http://www.eco-info.dk/dgb/detail.asp?id=#0#&count=1&sektion1=dgs&recid1=#0#&recname1=#1#&offset=#nr#"" target=""_blank"">#2#</a></b></font>#3#<br><br>')" & chr(13)
	endtext="document.write('<a href=""http://www.eco-info.dk"" target=""_blank""><font style=""color:#999999"">Leveret af Øko-info</font></a>')" & chr(13)

	SET rs= Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_ecoinfo_STRING

	theorgs=split(orgid,",")

	FOR i=0 to ubound(theorgs)
		rs.Source = "SELECT * FROM eisys_orgsettings WHERE orgid=" & theorgs(i)
		rs.Open()
	
		IF NOT rs.EOF THEN
			IF rs.Fields.Item("alldgb").value=1 THEN
				theSQL="SELECT DISTINCT om.id,om.name,m.title,m.shortdescription,m.id FROM (eilib_maindata m LEFT JOIN eilib_r_org o ON m.id=o.libid) LEFT JOIN eiorg_maindata om ON o.orgid=om.id WHERE om.id=" & theorgs(i) & " ORDER BY m.id DESC"
				filename="/log/ext/" & theorgs(i) & "_dgb_all.js"
				MakeIncFile theSQL,MM_ecoinfo_STRING,startText,endText,modelText,filename
			END IF
		END IF
		rs.close()
	next
	rs=""
END SUB

SUB MakeIncFile(theSQL,MM_ecoinfo_STRING,StartText,EndText,ModelText,Filename)
	DIM theRS, Format_Sel, fso, ts,theData,rowCount, colCount,i,ii, thisRow, theFile
	SET theRS= Server.CreateObject("ADODB.Recordset")
	theRS.ActiveConnection = MM_ecoinfo_STRING
	theRS.Source = theSQL
	theRS.Open()
	theData=theRS.GetRows
	theRS.Close()
	theRS=""
	colCount=uBound(theData,2)
	rowCount=uBound(theData,1)
	theFile=StartText
	FOR i=0 TO colCount
		thisRow=ModelText
		FOR ii=0 TO rowCount
			thisRow=replace(thisRow,"#" & ii & "#",trim(replace(replace(replace(theData(ii,i) & "","'","\'"),chr(13),"<br>"),chr(10),"")) & "")
		NEXT
		theFile=theFile & thisrow
	NEXT
	theFile=theFile & EndText
	thefile=replace(theFile,"<br><br><br>","<br><br>")
	WriteFile theFile,Filename
END SUB

%>