<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Function editForms()
Dim validate,bu,econet
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
	

	
Dim objConn,objRec,objRec3
Set objConn = Server.CreateObject ("ADODB.Connection") 
	objConn.Open MM_ecoinfo_STRING 
	'find artikel i tabel
Set objRec = Server.CreateObject ("ADODB.Recordset")
	objRec.Open "bu_Artikel", objConn, 0, 2, 2
	
	While not objRec.EOF 
		if CStr(objRec("Artikel_ID"))=CStr(request("artikelID")) then
		response.write(objRec("Artikel_ID"))
		response.end
		'opdater artikel
	'	objRec("Header")=request("title")
	'	objRec("Approved")=validate
	'	objRec("CreateDate")=request("createDate")
	'	objRec("ShortDescription")=request("shortdescription")
	'	objRec("FileName")=request("link")
	'	objRec("Content")=request("description")
	'	objRec.Update
		end if
	objRec.MoveNext
	Wend
objRec.Close
	
	'opdater site-info til tabel
Set objRec3 = Server.CreateObject ("ADODB.Recordset")

objRec3.Open "bu_Artikel_site", objConn, 0, 2, 2

	While not objRec3.EOF 
	if objRec3("artikel_ID")=request("artikelID") then
	objRec3("newwin")=CInt(request("hiddenwin"))	'0= nyt vindue, 1= samme vindue
	objRec3("busite")=bu							'1= vises på bu				
	objRec3("econetsite")=econet					'1= vises på econet
	objRec3.Update
	end if
	objRec3.MoveNext
	Wend
objRec3.Close
objConn.Close
End Function
%>