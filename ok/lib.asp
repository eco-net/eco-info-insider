<!--include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<!--#include virtual="/shared/insiderusers.asp" -->

<%
Dim theorgID,theID
SUB DoSave(act)
	'**** Creating SQL-statement
	IF act=1 THEN 'Insert
		DoCreate
		regUser Session("username"),theorgID,act,"ok",theID,request("title")
	ELSEIF act=2 THEN 'Update
		DoUpdate
		regUser Session("username"),theorgID,act,"ok",theID,request("title")
	ELSEIF act=3 THEN 'Delete
		DoDelete
	END IF

	'**** Showing confirmation to user
	response.redirect("confirmation.asp?done=" & act & request("params"))
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
	if request("econet")<>"" then
		econet=1 'vises p econet
		else
		econet=0
	end if
	status=2
	if (request.cookies("eiorgname") = "&Oslash;ko-net medarbejder") then 
		status=1
	end if

	'1. Inserting org.
	theSQL="INSERT INTO eical_maindata(" &_
	"title,subtitle,organizers,shortdescription,description,startdate,enddate,postnr,emailaddress,www,phone,validated,keywords,status,maker,starttime,endtime,place,address,econet,img_id,img_id2" &_
	") VALUES (" &_
	"'" & replace(request("title"),"'","''") & "'" & "," &_
	"'" & replace(request("subtitle"),"'","''") & "'" & "," &_
	"'" & replace(request("organizers"),"'","''") & "'" & "," &_
	"'" & replace(request("shortdescription"),"'","''") & "'" & "," &_
	"'" & replace(request("description"),"'","''") & "'" & "," &_
	FormatDateDB(request("startdate")) & "," &_
	FormatDateDB(request("enddate")) & "," &_
	request("postnr") & "," &_
	"'" & request("emailaddress") & "'" & "," &_
	"'" & request("www") & "'" & "," &_
	"'" & request("phone") & "'" & "," &_
	"0" & "," &_
	"'" & replace(request("keywords"),"'","''") & "'" & ","  &_
	status  & "," &_
	"'" & Session("username") & "'" & "," &_
	"'" & replace(request("starttime"),"'","''") & "'" & ","  &_
	"'" & replace(request("endtime"),"'","''") & "'" & ","  &_
	"'" & replace(request("place"),"'","''") & "'" & ","  &_
	"'" & replace(request("address"),"'","''") & "'" & "," &_
	econet & "," &_
	img_id & "," &_
	img_id2 &_
	")"
	'response.write theSQL
	'response.end
	theID=InsertRec(theSQL,"eical_maindata","id","title='" & replace(request("title"),"'","''") & "'")
	
	'2. Inserting Categories.
	theSQL=""
	FOR each Item in request("selCat")
		theSQL=theSQL & "INSERT INTO eical_r_cat (calid,catid) VALUES (" & theID & "," & item & ")"
	NEXT
	execCommand theSQL
	
	'3. Inserting Subjects
	theSQL=""
	FOR each Item in request("selsbj")
		theSQL=theSQL & "INSERT INTO eical_r_sbj (calid,sbjid) VALUES (" & theID & "," & item & ");" & chr(13)
	NEXT
	execCommand theSQL
	

	'4. Linking to Org
	theSQL=""
	theorgs=split(request("orgid"),",")
	theorgID=theorgs(0)'for regUser()
	FOR i=0 to ubound(theorgs)
		theSQL=theSQL & "INSERT INTO eical_r_org (orgid,calid) VALUES (" & theorgs(i) & "," & theID & ");" & chr(13)
	NEXT
	execCommand theSQL

	'5. Updating include-files
	if request("orgid")<>"" then
		'UpdateIncludes
	end if
	
	'6. Sender mail til org
	if LEN(request.cookies("eiuserid"))>0 then
		theBody=ReadFile("/admin/ei/snippets/okmail.txt")
		theBody=replace(theBody,"#titel#",request("title"))
		theBody=replace(theBody,"#id#",theID)
		for i=0 to ubound(theorgs)
			set rs=GetRecordSet("Select o.id as orgid, o.emailaddress1 as m, o.name, u.username,u.password from eiorg_maindata o inner join eisys_insiderusers u on o.id=u.orgid where o.id="&theorgs(i))
			tBody=replace(theBody,"#org#",rs("name"))
			tBody=replace(tBody,"#username#",rs("username"))
			tBody=replace(tBody,"#password#",rs("password"))
			tBody=replace(tBody,"#orgid#",rs("orgid"))
			SendCDOMail rs("m"),"eco-net@eco-net.dk","Vi har lagt Jeres arrangement ind",tBody
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
	if request("econet")<>"" then
		econet=1 'vises p econet
	else
		econet=0
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
	'fulladdress bevares hvis ikke den er tom:
	if request("fulladdress")<>"" then
	fa = replace(request("fulladdress"),"'","''")
	else
	fa = ""
	end if
	
	
	theSQL="UPDATE eical_maindata SET " &_
	"title='" & replace(request("title"),"'","''") & "'" & "," &_
	"subtitle='" & replace(request("subtitle"),"'","''") & "'" & "," &_
	"shortdescription='" & replace(request("shortdescription"),"'","''") & "'" & "," &_
	"description='" & replace(request("description"),"'","''") & "'" & "," &_
	"organizers='" & replace(request("organizers"),"'","''") & "'" & "," &_
	"startdate=" & FormatDateDB(request("startdate")) & "," &_
	"enddate=" & FormatDateDB(request("enddate")) & "," &_
	"postnr=" & request("postnr") & "," &_
	"emailaddress='" & request("emailaddress") & "'" & "," &_
	"www='" & request("www") & "'" & "," &_
	"phone='" & request("phone") & "'" & "," &_
	"validated=" & validated & "," &_
	"modidate=" & FormatDateDB(date) & "," &_
	"keywords='" & replace(request("keywords"),"'","''") & "'" & ","  &_
	"status=" & status & "," &_
	"editor='" & Session("username") & "'" & "," &_
	"fulladdress='" & fa & "'" & "," &_
	"starttime='" & replace(request("starttime"),"'","''") & "'" & "," &_
	"endtime='" & replace(request("endtime"),"'","''") & "'" & "," &_
	"place='" & replace(request("place"),"'","''") & "'" & "," &_
	"address='" & replace(request("address"),"'","''") & "'" & "," &_
	"econet=" & econet & "," &_
	"img_id=" & img_id & "," &_
	"img_id2=" & img_id2 &_
	" WHERE id=" & request("id")
	
	execCommand theSQL
	
	
	'2. Updating categories
	IF LEN(request("selCat"))>0 THEN
		theSQL="DELETE FROM eical_r_cat WHERE calid=" & request("id")
		execCommand theSQL
		theSQL=""
		FOR each Item in request("selCat")
			theSQL=theSQL & "INSERT INTO eical_r_cat (calid,catid) VALUES (" & request("id") & "," & item & ");" & chr(13)
		NEXT
		'response.write theSQL & "<br><bR>"
		execCommand theSQL
	END IF

	'3. Updating subjects
	IF LEN(request("selSbj"))>0 THEN
		theSQL="DELETE FROM eical_r_sbj WHERE calid=" & request("id")
		execCommand theSQL
		theSQL=""
		FOR each Item in request("selsbj")
			theSQL=theSQL & "INSERT INTO eical_r_sbj (calid,sbjid) VALUES (" & request("id") & "," & item & ");" & chr(13)
		NEXT
		'response.write theSQL & "<br><br>"
		execCommand theSQL
	END IF

	'4. Updating organizations
	IF len(request("theorgid"))>0 THEN
		theSQL = "DELETE FROM eical_r_org WHERE calid=" & request("id")
		execCommand theSQL
		response.write theSQL & "<br>"
		theSQL = ""
		theorgs=split(request("theorgid"),",")
		
		FOR i=0 to ubound(theorgs)
			if LEN(theorgs(i))>1 then theSQL=theSQL & "INSERT INTO eical_r_org (calid,orgid) VALUES (" & request("id") & "," & theorgs(i) & ");" & chr(13)
		NEXT
		execCommand theSQL
	END IF

	'5. Updating include-files
	if request("orgid")<>"" OR request("theorgid")<>"" then
		UpdateIncludes
	end if
	theorgID=request("firstorgid")'for regUser()
END SUB

SUB DoDelete
	theid=request("id")
	theSQL="Delete FROM eical_r_sbj WHERE calid=" & theid
	execCommand theSQL
	theSQL="Delete FROM eical_r_cat WHERE calid=" & theid
	execCommand theSQL
	theSQL="Delete FROM eical_r_coll WHERE calid=" & theid
	execCommand theSQL
	theSQL="Delete FROM eical_r_org WHERE calid=" & theid
	execCommand theSQL
	theSQL="Delete FROM eical_maindata WHERE id=" & theid
	execCommand theSQL

	'4. Updating include-files
	'IF LEN(request("orgid"))>0 THEN UpdateIncludes
END SUB

SUB UpdateIncludes

	starttext=""
	modelText="document.write('<font style=""font-size:larger""><b><a href=""http://www.eco-info.dk/ok/detail.asp?id=#0#&count=1&sektion1=dgs&recid1=#0#&recname1=#1#&offset=#nr#"" target=""_blank"">#2#</a></b></font><br>#3##4#<br><br>')" & chr(13)
	endtext="document.write('<a href=""http://www.eco-info.dk"" target=""_blank""><font style=""color: #999999"">Leveret af ko-info</font></a>')" & chr(13)

	SET rs= Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_ecoinfo_STRING
	if request("orgid")<>"" then
		theorgs=split(request("orgid"),",")
	else
		theorgs=split(request("theorgid"),",")
	end if
	FOR i=0 to ubound(theorgs)
		IF len(theorgs(i))>1 then
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
		end if
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
'	response.end
END SUB
%>