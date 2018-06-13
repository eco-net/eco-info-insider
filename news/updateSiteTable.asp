<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Set objConn = Server.CreateObject ("ADODB.Connection") 
	objConn.Open MM_ecoinfo_STRING 
	
	Set objRec2 = Server.CreateObject ("ADODB.Recordset")
	objRec2.Open "bu_Artikel", objConn, 1, 1, 2
	Set objRec3 = Server.CreateObject ("ADODB.Recordset")
	objRec3.Open "bu_Artikel_site", objConn, 0, 2, 2
	
While not objRec2.EOF
	objRec3.AddNew
	objRec3("artikel_id")=objRec2("Artikel_ID")
	objRec3("newwin")=0						'0= nyt vindue, 1= samme vindue
	objRec3("busite")=1							'1= vises på bu				
	objRec3("econetsite")=0					'1= vises på econet
	objRec2.MoveNext
Wend
objRec3.Update
objRec2.Close
objRec3.Close
objConn.Close
%>