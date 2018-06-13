<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include file="homepage.asp" -->

<%
SUB DoSave(thefilename,lang,act)
	'**** Creating SQL-statement
	DIM translated
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
		DeleteFileInDB "hw2_splashes","imagepath","id",request("id")
		theSQL="Delete From hw2_splashes WHERE id=" & request("id")
		execCommand theSQL
		IF act=4 THEN WriteSplashes lang
	ELSEIF act=5 THEN 'Choose
		DoChoose lang
	END IF

	IF act<5 THEN
		UpdateMenuFiles lang,translated
	END IF	

	'**** Showing confirmation to user
	response.redirect("confirmation.asp?done=" & act)
END SUB

SUB DoCreate(thefilename,thelang,translated)
	theSQL="INSERT INTO hw2_splashes(" &_
		"headline,content,link,imagepath,language,linkaction,translated" &_
		") VALUES (" &_
		"'" & replace(UploadRequest.Item("headline").Item("Value") & "","'","''") & "'" & "," &_
		"'" & replace(UploadRequest.Item("content").Item("Value") & "","'","''") & "'" & "," &_
		"'" & replace(UploadRequest.Item("pagelink").Item("Value") & "","'","''") & "'" & "," &_
		"'" & thefilename & "'," &_
		"'" & thelang & "'," &_
		"'" & replace(UploadRequest.Item("linkaction").Item("Value") & "","'","''") & "'," &_
		translated & ")"
	execCommand theSQL
END SUB

SUB DoUpdate(thefilename,lang)
	'Deleting image
	IF thefilename<>"" THEN DeleteFileInDB "hw2_splashes","imagepath","id",UploadRequest.Item("id").Item("Value")

	theSQL="UPDATE hw2_splashes SET " &_
	"headline='" & UploadRequest.Item("headline").Item("Value") & "'" & "," &_
	"content='" & UploadRequest.Item("content").Item("Value") & "'" & "," &_
	"link='" & replace(UploadRequest.Item("pagelink").Item("Value"),"'","''") & "'," &_
	"linkaction='" & replace(UploadRequest.Item("linkaction").Item("Value"),"'","''") & "'"
	IF thefilename<>"" THEN theSQL=theSQL & ",imagepath='" & thefilename & "'"
	IF UploadRequest.Exists("translated") THEN 
		IF LEN(UploadRequest.Item("translated").Item("Value"))>0 THEN theSQL=theSQL & ",translated=1"
	END IF
	theSQL=theSQL & " WHERE id=" & UploadRequest.Item("id").Item("Value")
	execCommand theSQL

	'Update Homepage, if applicable
	Set rs = Server.CreateObject("ADODB.recordset")
	rs.ActiveConnection=MM_data_STRING
	rs.source="SELECT id FROM hw2_pageelements WHERE language='" & lang & "' AND category=6 AND content_number=" & UploadRequest.Item("id").Item("Value")
	rs.open()
	IF NOT rs.EOF THEN writeSplashes lang
	rs.close()
	rs=""
END SUB

SUB DoChoose(lang)
	' 1. deleting exsisting
	theSQL="DELETE FROM hw2_pageelements WHERE category=6 AND language='" & lang & "'"
	execCommand theSQL
	' 2. creating new
	theSQL=""
	theCount=0
	IF CINT(request("spl1"))>0 THEN 
		theSQL=theSQL & "INSERT INTO hw2_pageelements (pageid,category,content_number,number,language) VALUES (1,6," & request("spl1") & ",1,'" & lang & "');"
		theCount=theCount+1
	END IF
	IF CINT(request("spl2"))>0 THEN 
		theSQL=theSQL & "INSERT INTO hw2_pageelements (pageid,category,content_number,number,language) VALUES (1,6," & request("spl2") & ",2,'" & lang & "');"
		theCount=theCount+1
	END IF
	IF CINT(request("spl3"))>0 THEN 
		theSQL=theSQL & "INSERT INTO hw2_pageelements (pageid,category,content_number,number,language) VALUES (1,6," & request("spl3") & ",3,'" & lang & "');"
		theCount=theCount+1
	END IF
	execCommand theSQL
	WriteSplashes lang
END SUB

SUB UpdateMenuFiles(lang,translated)
	MakeMenuFile "SELECT id,headline,translated FROM hw2_splashes WHERE language='" & lang & "' ORDER BY 2","&lt;Select&gt;","id","headline","/backend/includes/menufiles/sel_splashes_dyn_" & lang & ".txt","cursel",translated
	MakeMenuFile "SELECT id,headline,translated FROM hw2_splashes WHERE language='" & lang & "' ORDER BY 2","&lt;Select&gt;","id","headline","/backend/includes/menufiles/sel_splashes_" & lang & ".txt","",translated
END SUB
%>