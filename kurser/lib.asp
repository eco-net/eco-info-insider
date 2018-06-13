<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<!--#include virtual="/shared/insiderusers.asp" -->
<%setlocale(1030) %>
<%
Dim theorgID,theID
theorgID=request("orgid")
SUB DoSave(act)
	'**** Creating SQL-statement
	IF act=1 THEN 'Insert
		DoCreate
		regUser Session("username"),theorgID,act,"kurser",theID,request("title")
	ELSEIF act=2 THEN 'Update
		DoUpdate
		regUser Session("username"),theorgID,act,"kurser",request("id"),request("title")
	ELSEIF act=3 THEN 'Delete
		DoDelete
	END IF

	'**** Showing confirmation to user
	response.redirect("confirmation.asp?done=" & act & request("params"))
END SUB

SUB DoCreate

status=2
creator=2
if (request.cookies("eiorgname") = "&Oslash;ko-net medarbejder") then 
status=1
creator=1
end if
	'1. Inserting org.
	theSQL="INSERT INTO eicourse_maindata(" &_
	"title,subtitle,organizers,shortdescription,description,education,startdate,enddate,starttime,hometime,school,address,postnr,emailadress,www,phone,validated,keywords,status,creator,maker" &_
	") VALUES (" &_
	"'" & replace(request("title"),"'","''") & "'" & "," &_
	"'" & replace(request("subtitle"),"'","''") & "'" & "," &_
	"'" & replace(request("organizers"),"'","''") & "'" & "," &_
	"'" & replace(request("shortdescription"),"'","''") & "'" & "," &_
	"'" & replace(request("description"),"'","''") & "'" & "," &_
	"'" & replace(request("education"),"'","''") & "'" & "," &_
	FormatDateDB(request("startdate")) & "," &_
	FormatDateDB(request("enddate")) & "," &_
	"'" & replace(request("starttime"),"'","''") & "'" & "," &_
	"'" & replace(request("hometime"),"'","''") & "'" & "," &_
	"'" & replace(request("school"),"'","''") & "'" & "," &_
	"'" & replace(request("address"),"'","''") & "'" & "," &_
	request("postnr") & "," &_
	"'" & request("emailaddress") & "'" & "," &_
	"'" & request("www") & "'" & "," &_
	"'" & request("phone") & "'" & "," &_
	"0" & "," &_
	"'" & replace(request("keywords"),"'","''") & "'" & ","  &_
	status & "," &_
	creator & "," &_
	"'" & Session("username") & "'" &_
	")"

	theID=InsertRec(theSQL,"eicourse_maindata","id","title='" & replace(request("title"),"'","''") & "'")
	
	'2. Inserting Categories.
	theSQL=""
	FOR each Item in request("selCat")
		theSQL=theSQL & "INSERT INTO eicourse_r_cat (courseid,catid) VALUES (" & theID & "," & item & ")"
	NEXT
	execCommand theSQL
	
	'3. Inserting Subjects
	theSQL=""
	FOR each Item in request("selsbj")
		theSQL=theSQL & "INSERT INTO eicourse_r_sbj (courseid,sbjid) VALUES (" & theID & "," & item & ");" & chr(13)
	NEXT
	execCommand theSQL

	'4. Linking to Org
	theSQL=""
	theorgs=split(request("orgid"),",")
	FOR i=0 to ubound(theorgs)
		theSQL=theSQL & "INSERT INTO eicourse_r_org (orgid,courseid) VALUES (" & theorgs(i) & "," & theID & ")"
	NEXT
	execCommand theSQL

	'5. Updating include-files
'	if request("orgid")<>"" then
'		UpdateIncludes
'	end if
END SUB

SUB DoUpdate

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
	
	theSQL="UPDATE eicourse_maindata SET " &_
	"title='" & replace(request("title"),"'","''") & "'" & "," &_
	"subtitle='" & replace(request("subtitle"),"'","''") & "'" & "," &_
	"shortdescription='" & replace(request("shortdescription"),"'","''") & "'" & "," &_
	"description='" & replace(request("description"),"'","''") & "'" & "," &_
	"education='" & replace(request("education"),"'","''") & "'" & "," &_
	"organizers='" & replace(request("organizers"),"'","''") & "'" & "," &_
	"startdate=" & FormatDateDB(request("startdate")) & "," &_
	"enddate=" & FormatDateDB(request("enddate")) & "," &_
	"starttime='" & request("meettime") & "'" & "," &_
	"hometime='" & request("hometime") & "'" & "," &_
	"school='" & request("school") & "'" & "," &_
	"address='" & replace(request("address"),"'","''") & "'" & "," &_
	"postnr=" & request("postnr") & "," &_
	"emailadress='" & request("emailaddress") & "'" & "," &_
	"www='" & request("www") & "'" & "," &_
	"phone='" & request("phone") & "'" & "," &_
	"validated=" & validated & "," &_
	"modidate=" & FormatDateDB(date) & "," &_
	"keywords='" & replace(request("keywords"),"'","''") & "'" & ","  &_
	"status=" & status & "," &_
	"editor='" & Session("username") & "'" &_
	" WHERE id=" & request("id")
	
	execCommand theSQL

	'2. Updating categories
	IF LEN(request("selCat"))>0 THEN
		theSQL="DELETE FROM eicourse_r_cat WHERE courseid=" & request("id")
		execCommand theSQL
		theSQL=""
		FOR each Item in request("selCat")
			theSQL=theSQL & "INSERT INTO eicourse_r_cat (courseid,catid) VALUES (" & request("id") & "," & item & ");" & chr(13)
		NEXT
		'response.write theSQL & "<br><bR>"
		execCommand theSQL
	END IF

	'3. Updating subjects
	IF LEN(request("selSbj"))>0 THEN
		theSQL="DELETE FROM eicourse_r_sbj WHERE courseid=" & request("id")
		execCommand theSQL
		theSQL=""
		FOR each Item in request("selsbj")
			theSQL=theSQL & "INSERT INTO eicourse_r_sbj (courseid,sbjid) VALUES (" & request("id") & "," & item & ");" & chr(13)
		NEXT
		'response.write theSQL & "<br><br>"
		execCommand theSQL
	END IF

	'4. Updating organizations
	'IF LEN(request("orgid"))>0 THEN
	''theSQL="DELETE FROM eicourse_r_org WHERE courseid=" & request("id")
	'	execCommand theSQL
	'	theSQL=""
	'	theorgs=split(request("theorgid"),",")
	'	response.write(theorgs)
	'	response.end
	'	FOR i=0 to ubound(theorgs)
	'		theSQL=theSQL & "INSERT INTO eicourse_r_org (courseid,orgid) VALUES (" & request("id") & "," & theorgs(i) & ");" & chr(13)
	'	NEXT
		'response.write theSQL & "<br><br>"
	'	execCommand theSQL
	'END IF

	'5. Updating include-files
	'if request("orgid")<>"" then
	'	UpdateIncludes
	'end if
	'response.end
END SUB

SUB DoDelete

	theid=request("id")
	
	theSQL="Delete FROM eicourse_r_sbj WHERE courseid=" & theid
	execCommand theSQL
	
	theSQL="Delete FROM eicourse_r_cat WHERE courseid=" & theid
	execCommand theSQL
	
	theSQL="Delete FROM eicourse_r_org WHERE courseid=" & theid
	execCommand theSQL
	
	theSQL="Delete FROM eicourse_maindata WHERE id=" & theid
	execCommand theSQL

	'4. Updating include-files
	'IF LEN(request("orgid"))>0 THEN UpdateIncludes
END SUB

SUB UpdateIncludes
	starttext=""
	modelText="document.write('<font style=""font-size:larger""><b><a href=""http://www.oko-info.dk/ok/detail.asp?id=#0#&count=1&sektion1=dgs&recid1=#0#&recname1=#1#&offset=#nr#"" target=""_blank"">#2#</a></b></font><br>#3##4#<br><br>')" & chr(13)
	endtext="document.write('<font style=""font-size:smaller""><a href=""http://www.oko-info.dk"" target=""_blank"">Leveret af Øko-info</a></font>')" & chr(13)

	SET rs= Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_ecoinfo_STRING
	theorgs=split(request("orgid"),",")
	FOR i=0 to ubound(theorgs)
		rs.Source = "SELECT * FROM eisys_orgsettings WHERE orgid=" & theorgs(i)
		rs.Open()
	
		IF NOT rs.EOF THEN
			IF rs.Fields.Item("allOk").value=1 THEN
				theSQL="SELECT DISTINCT om.id,om.name,m.title,m.startdate,m.enddate,p.city FROM ((eical_maindata m LEFT JOIN eisys_postnr p ON m.postnr=p.postnr) LEFT JOIN eical_r_org o ON m.id=o.calid) LEFT JOIN eiorg_maindata om ON o.orgid=om.id WHERE om.id=" & theorgs(i) & " ORDER BY m.startdate DESC"
				filename="/log/ext/" & theorgs(i) & "_ok_all.js"
				MakeIncFileOK theSQL,startText,endText,modelText,filename
			END IF
		
			IF rs.Fields.Item("newOk").value=1 THEN
				theSQL="SELECT DISTINCT om.id,om.name,m.title,m.startdate,m.enddate,p.city FROM ((eical_maindata m LEFT JOIN eisys_postnr p ON m.postnr=p.postnr) LEFT JOIN eical_r_org o ON m.id=o.calid) LEFT JOIN eiorg_maindata om ON o.orgid=om.id WHERE om.id=" & theorgs(i) & " AND m.startdate>=" & FormatDateDB(Date) & " ORDER BY m.startdate ASC"
				filename="/log/ext/" & theorgs(i) & "_ok_new.js"
				MakeIncFileOK theSQL,startText,endText,modelText,filename
			END IF
		END IF
		rs.close()
	NEXT
	rs=""
END SUB

SUB MakeIncFileOK(theSQL,startText,endText,modelText,filename)
	SET theRS= Server.CreateObject("ADODB.Recordset")
	theRS.ActiveConnection = MM_ecoinfo_STRING
	theRS.Source = theSQL
	theRS.Open()
	theFile=""
	i=0
	WHILE NOT theRS.EOF
		thisRow=ModelText
		thisRow=replace(thisRow,"#0#",theRS.Fields.Item("id").value & "")
		thisRow=replace(thisRow,"#1#",replace(replace(replace(theRS.Fields.Item("name").value,"'","\'"),chr(13),"<br>"),chr(10),"") & "")
		thisRow=replace(thisRow,"#2#",replace(replace(replace(theRS.Fields.Item("title").value,"'","\'"),chr(13),"<br>"),chr(10),"") & "")
		IF DatePart("m",theRS.Fields.Item("startdate").value)=DatePart("m",theRS.Fields.Item("enddate").value) THEN
			IF DatePart("d",theRS.Fields.Item("startdate").value)=DatePart("d",theRS.Fields.Item("enddate").value) THEN
				thisRow=replace(thisRow,"#3#",FormatDateTime(theRS.Fields.Item("startdate").value,vbLongDate))
			ELSE
				thisRow=replace(thisRow,"#3#",DatePart("d",theRS.Fields.Item("startdate").value) & " til " & FormatDateTime(theRS.Fields.Item("enddate").value,vbLongDate))
			END IF
		ELSEIF theRS.Fields.Item("enddate").value<>"" THEN
			thisRow=replace(thisRow,"#3#",FormatDateTime(theRS.Fields.Item("startdate").value,vbLongDate) & " til " & FormatDateTime(theRS.Fields.Item("enddate").value,vbLongDate))
		ELSE
			thisRow=replace(thisRow,"#3#",FormatDateTime(theRS.Fields.Item("startdate").value,vbLongDate) & "bla")
		END IF	
	
		IF theRS.Fields.Item("city").value<>"" THEN
			thisRow=replace(thisRow,"#4#"," | " & theRS.Fields.Item("city").value)
		ELSE
			thisRow=replace(thisRow,"#4#","")
		END IF
		theFile=theFile & replace(thisRow,"#nr#",i)
		theRS.movenext
		i=i+1
	WEND
	IF theFile="" THEN
		theFile="document.write('Der er ingen arrangementer.')"
	ELSE
		theFile=startText & theFile & EndText
	END IF

	WriteFile theFile,Filename
	theRS.close()
	theRS=""
END SUB
%>