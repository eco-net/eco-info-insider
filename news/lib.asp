<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp" -->

<%

SUB DoUpdate
	if request("sletNyhed")="slet" then
	DoDelete
	else
	DoSave
	end if
END SUB
SUB DoDelete
	Dim ID, objConn, objRec
	ID=CInt(request("id"))
	Set objConn = Server.CreateObject ("ADODB.Connection") 
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	objConn.Open MM_ecoinfo_STRING 
	objRec.Open "bu_Artikel", objConn, 0, 2, 2
	While Not objRec.EOF
		if objRec("Artikel_ID")=ID then
			objRec.Delete
			objRec.Update
		end if
	objRec.MoveNext
	wend
	objRec.Close
	objRec.Open "bu_Artikel_site", objConn, 0, 2, 2
	While Not objRec.EOF
		if objRec("artikel_ID")=ID then
			objRec.Delete
			objRec.Update
		end if
	objRec.MoveNext
	wend
	objRec.Close
	objConn.Close
	response.redirect "index.asp"
END SUB
SUB DoSave

	Dim validate,bu,econet,newwin,img
	if request("img_id")<>"" then
	img=CInt(request("img_id"))
	else
	img=0
	end if
	if request("img_id2")<>"" then
	img2=CInt(request("img_id2"))
	else
	img2=0
	end if
	if request("validated")<>"" then
	validate=1 	'valideret
	else 
	validate=0
	end if
	if request("bu")<>"" then
	bu=1 ' vises på bu
	else
	bu=0
	end if
	if request("econet")<>"" then
	econet=1 'vises på econet
	else
	econet=0
	end if
	if request("vindue")=0 then
	newwin=0
	else
	newwin=1
	end if
	if request("miljoinfo")<>"" then
	miljoinfo=1 'Miljø Info tekst vises
	else
	miljoinfo=0
	end if
	cat=CInt(request("cat"))
	
	'1. Updating Nyhed
	theSQL="UPDATE bu_Artikel SET " &_
	"Header='" & toHTML(request("title")) & "'" & "," &_
	"Approved=" & validate & "," &_
	"ShortDescription='" & toHTML(request("shortdescription")) & "'" & "," &_
	"FileName='" & request("link") & "'" & "," &_
	"FileName2='" & request("link2") & "'" & "," &_
	"FileName3='" & request("link3") & "'" & "," &_
	"link4='" & request("link4") & "'" & "," &_
	"Content='" & toHTML(request("EditorValue")) & "'" & "," &_
	"img_id=" & img & "," &_
	"ArtikelCategory_ID=" & cat & "," &_
	"linktext='" & request("linktext") & "'" & "," &_
	"linktext2='" & request("linktext2") & "'" & "," &_
	"linktext3='" & request("linktext3") & "'" & "," &_
	"linktext4='" & request("linktext4") & "'" & "," &_
	"img_text='" & request("img_text") & "'" & "," &_
	"img_text2='" & request("img_text2") & "'" & "," &_
	"Tema_ID=" & request("tema")  & "," &_
	"img_id2=" & img2 & "," &_
	"imagefilename1='" & request("imagefilename1") & "'," &_
	"imagefilename2='" & request("imagefilename2") & "'," &_
	"miljoinfo=" & miljoinfo &_
	" WHERE Artikel_ID=" & request("id")
	'response.write theSQL
	'response.end
	execCommand theSQL

	'2. Updating site-info
	theSQL="UPDATE bu_Artikel_site SET " &_
	"newwin=" & newwin & "," &_
	"busite=" & bu & "," &_
	"econetsite=" & econet &_
	" WHERE artikel_ID=" & request("id")
	execCommand theSQL
END SUB
FUNCTION FormateDateDB(theDate)
	IF LEN(theDate)>0 THEN
		theDate=CDate(theDate)
		FormateDateDB="{ts '" & datepart("yyyy",theDate) & "-" & right("0" & CStr(datepart("m",theDate)),2) & "-" &_
			right("0" & CSTR(datepart("d",theDate)),2) & " 00:00:00'}"
	ELSE
		FormateDateDB="''"
	END IF
END FUNCTION
%>