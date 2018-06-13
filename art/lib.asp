<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->
<%
	

Dim newid 'Nyoprettet artikel returnerer id til brug for videresendelse til edit.asp
%>
<%
SUB DoSave(act)

	'**** Creating SQL-statement
	IF act=1 THEN 'Insert
		DoCreate
	ELSEIF act=2 THEN 'Update
		DoUpdate
	ELSEIF act=3 THEN 'Delete
		DoDelete
	END IF
	if request("saveandreturn")<>"" then 'videresendelse til edit.asp
		if act=2 then
			response.redirect "edit.asp?id=" & request("id")
		elseif act=1 then
			response.redirect "edit.asp?id=" & newid
		end if
	else
	'**** Showing confirmation to user
	response.redirect "confirmation.asp?done=" & act & request("params")
	response.end
	end if
END SUB

SUB DoCreate



	'1. Inserting org.
	
	theSQL="INSERT INTO eiart_maindata(" &_
	"title,subtitle,shortdescr,descr,lang,author,authordate,editor,editormail,editorwww,webaddress,validated,keywords,maker" &_
	") VALUES (" &_
	
	"'" & replace(request("title"),"'","''") & "'" & "," &_
	"'" & replace(request("subtitle"),"'","''") & "'" & "," &_
	"'" & replace(request("shortdescription"),"'","''") & "'" & "," &_
	"'" & replace(request("description"),"'","''") & "'" & "," &_
	"'" & request("language") & "'" & "," &_
	"'" & replace(request("author"),"'","''") & "'" & "," &_
	"'" & request("year") & "'" & "," &_
	"'" & replace(request("editor"),"'","''") & "'" & "," &_
	"'" & request("editoremail") & "'" & "," &_
	"'" & request("editorwww") & "'" & "," &_
	"'" & request("webaddress") & "'" & "," &_
	"0" & "," &_
	"'" & replace(request("keywords"),"'","''") & "'"  & ","  &_
	"'" & Session("username") & "'" &_
	")"
			
	


	theID=InsertRec(theSQL,"eiart_maindata","id","")
	newid=theID
	
	'2. Inserting Categories.
	theSQL="INSERT INTO eiart_r_cat (artid,catid) VALUES (" & theID & "," & request("selCat") & ")"
	execCommand theSQL
	
	'3. Inserting Subjects
	theSQL=""
	FOR each Item in request("selsbj")
		theSQL=theSQL & "INSERT INTO eiart_r_sbj (artid,sbjid) VALUES (" & theID & "," & item & ");" & chr(13)
	NEXT
	execCommand theSQL

	'4. Linking to Org
	theSQL="INSERT INTO eiart_r_org (orgid,artid) VALUES (" & request("orgid") & "," & theID & ")"
	execCommand theSQL

	
	'5. Updating include-files
	'UpdateIncludes
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
	
	theSQL="UPDATE eiart_maindata SET " &_
	"title='" & replace(request("title"),"'","''") & "'" & "," &_
	"subtitle='" & replace(request("subtitle"),"'","''") & "'" & "," &_
	"shortdescr='" & replace(request("shortdescription"),"'","''") & "'" & "," &_
	"descr='" & replace(request("description"),"'","''") & "'" & "," &_
	"lang='" & request("language") & "'" & "," &_
	"author='" & replace(request("author"),"'","''") & "'" & "," &_
	"authordate='" & request("year") & "'" & "," &_
	"editor='" & replace(request("editor"),"'","''") & "'" & "," &_
	"editormail='" & request("editoremail") & "'" & "," &_
	"editorwww='" & request("editorwww") & "'" & "," &_
	"webaddress='" & request("webaddress") & "'" & "," &_
	"validated=" & validated & "," &_
	"modidate=" & FormatDateDB(date) & "," &_
	"keywords='" & replace(request("keywords"),"'","''") & "'" & "," &_
	"edit='" & Session("username") & "'" &_
	" WHERE id=" & request("id")
	
	execCommand theSQL
	'2. Updating categories
	IF LEN(request("selCat"))>0 THEN
		theSQL="UPDATE eiart_r_cat SET catid=" &  request("selCat") & " WHERE artid=" & request("id")
		execCommand theSQL
	END IF

	'3. Updating subjects
	IF LEN(request("selSbj"))>0 THEN
		theSQL="DELETE FROM eiart_r_sbj WHERE artid=" & request("id")
		execCommand theSQL
		theSQL=""
		FOR each Item in request("selsbj")
			theSQL=theSQL & "INSERT INTO eiart_r_sbj (artid,sbjid) VALUES (" & request("id") & "," & item & ");" & chr(13)
		NEXT
		execCommand theSQL
	END IF

	'4. Updating include-files
	'UpdateIncludes

END SUB

SUB DoDelete
	theid=request("id")
	theSQL="Delete FROM eiart_r_sbj WHERE artid=" & theid
	execCommand theSQL
	theSQL="Delete FROM eiart_r_cat WHERE artid=" & theid
	execCommand theSQL
	'theSQL="Delete FROM eiart_r_coll WHERE artid=" & theid
	'execCommand theSQL
	theSQL="Delete FROM eiart_r_org WHERE artid=" & theid
	execCommand theSQL
	theSQL="Delete FROM eiart_maindata WHERE id=" & theid
	execCommand theSQL

	'4. Updating include-files
	'IF LEN(request("orgid"))>0 THEN UpdateIncludes
END SUB

SUB UpdateIncludes
	starttext=""
	modelText="document.write('<font style=""font-size:larger""><b><a href=""http://www.eco-info.dk/dgb/detail.asp?id=#0#&count=1&sektion1=dgs&recid1=#0#&recname1=#1#&offset=#nr#"" target=""_blank"">#2#</a></b></font>#3#<br><br>')" & chr(13)
	endtext="document.write('<a href=""http://www.eco-info.dk"" target=""_blank""><font style=""color:#999999"">Leveret af Øko-info</font></a>')" & chr(13)

	SET rs= Server.CreateObject("ADODB.Recordset")
	rs.ActiveConnection = MM_ecoinfo_STRING
	rs.Source = "SELECT * FROM eisys_orgsettings WHERE orgid=" & request("orgid")
	rs.Open()

	IF NOT rs.EOF THEN
		IF rs.Fields.Item("alldgb").value=1 THEN
			theSQL="SELECT DISTINCT om.id,om.name,m.title,m.shortdescription,m.id FROM (eilib_maindata m LEFT JOIN eilib_r_org o ON m.id=o.libid) LEFT JOIN eiorg_maindata om ON o.orgid=om.id WHERE om.id=" & request("orgid") & " ORDER BY m.id DESC"
			filename="/log/ext/" & request("orgid") & "_dgb_all.js"
			MakeIncFile theSQL,MM_ecoinfo_STRING,startText,endText,modelText,filename
		END IF
	END IF
	rs.close()
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