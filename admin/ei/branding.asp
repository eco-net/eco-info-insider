<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include file="homepage.asp" -->

<%
SUB DoSave(thefilename,lang,act)
	'**** Creating SQL-statement
	DIM translated,thefilename2
	translated=0
	IF lang<>"dk" THEN translated=1

	IF act=1 THEN 'Insert
		DoCreate thefilename,lang,"1"
		
		'Copying to language-versions
		IF lang="dk" THEN
			IF UploadRequest.Exists("copyto_d") THEN 
				CopyImage thefilename,"d"
				DoCreate replace(thefilename,"dk","d"),"d","0"
				UpdateMenuFiles "d",1
			END IF
			IF UploadRequest.Exists("copyto_gb") THEN 
				CopyImage thefilename,"gb"
				DoCreate replace(thefilename,"dk","gb"),"gb","0"
				UpdateMenuFiles "gb",1
			END IF
		END IF

	ELSEIF act=2 THEN 'Update
		DoUpdate thefilename,lang
	ELSEIF act=3 OR act=4 THEN 'Delete
		DeleteFileInDB "hw2_branding","imagepath","id",request("id")
		theSQL="Delete From hw2_branding WHERE id=" & request("id")
		execCommand theSQL
		IF act=4 THEN WriteBa lang
	ELSEIF act=5 THEN 'Choose
		DoChoose lang
	END IF

	IF act<5 THEN
		UpdateMenuFiles lang,translated
	END IF	

	'**** Showing confirmation to user
	response.redirect("confirmation.asp?done=" & act)
END SUB

SUB DoCreate (fname,thelang,translated)
	theSQL="INSERT INTO hw2_branding(" &_
		"headline,content,link,imagepath,language,linkaction,translated" &_
		") VALUES (" &_
		"'" & replace(UploadRequest.Item("headline").Item("Value") & "","'","''") & "'" & "," &_
		"'" & replace(UploadRequest.Item("content").Item("Value") & "","'","''") & "'" & "," &_
		"'" & replace(UploadRequest.Item("pagelink").Item("Value") & "","'","''") & "'" & "," &_
		"'" & fname & "'," &_
		"'" & thelang & "'," &_
		"'" & replace(UploadRequest.Item("linkaction").Item("Value") & "","'","''") & "'," &_
		translated & ")"
	execCommand theSQL
END SUB

SUB DoUpdate (fname,thelang)
	'Deleting image
	IF fname<>"" THEN DeleteFileInDB "hw2_branding","imagepath","id",UploadRequest.Item("id").Item("Value")

	theSQL="UPDATE hw2_branding SET " &_
	"headline='" & UploadRequest.Item("headline").Item("Value") & "'" & "," &_
	"content='" & UploadRequest.Item("content").Item("Value") & "'" & "," &_
	"link='" & replace(UploadRequest.Item("pagelink").Item("Value"),"'","''") & "'," &_
	"linkaction='" & replace(UploadRequest.Item("linkaction").Item("Value"),"'","''") & "'"
	IF fname<>"" THEN theSQL=theSQL & ",imagepath='" & fname & "'"
	IF UploadRequest.Exists("translated") THEN 
		IF LEN(UploadRequest.Item("translated").Item("Value"))>0 THEN theSQL=theSQL & ",translated=1"
	END IF
	theSQL=theSQL & " WHERE id=" & UploadRequest.Item("id").Item("Value")
	execCommand theSQL
	'Update Homepage, if applicable
	Set rs = Server.CreateObject("ADODB.recordset")
	rs.ActiveConnection=MM_data_STRING
	rs.source="SELECT id FROM hw2_pageelements WHERE language='" & lang & "' AND category=7 AND content_number=" & UploadRequest.Item("id").Item("Value")
	rs.open()
	IF NOT rs.EOF THEN writeBa lang
	rs.close()
	rs=""
END SUB

SUB DoChoose(lang)
	' 1. deleting exsisting
	theSQL="DELETE FROM hw2_pageelements WHERE category=7 AND language='" & lang & "'"
	execCommand theSQL
	' 2. creating new
	theSQL=""
	thevar=split(request("sel"),",")
	i=1
	FOR each Item in thevar
		theSQL=theSQL & "INSERT INTO hw2_pageelements (pageid,category,content_number,language,number) VALUES (1,7," & Item & ",'" & lang & "'," & i & ");"
		i=i+1
	NEXT
	execCommand theSQL
	WriteBa lang
END SUB

SUB UpdateMenuFiles(lang,translated)
	MakeMenuFile "SELECT id,headline,translated FROM hw2_branding WHERE language='" & lang & "' ORDER BY 2","&lt;Select&gt;","id","headline","/backend/includes/menufiles/sel_branding_dyn_" & lang & ".txt","cursel",translated
	MakeMenuFile "SELECT id,headline,translated FROM hw2_branding WHERE language='" & lang & "' ORDER BY 2","&lt;Select&gt;","id","headline","/backend/includes/menufiles/sel_branding_" & lang & ".txt","",translated
END SUB
%>