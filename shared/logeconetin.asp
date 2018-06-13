<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
IF LEN(request("username"))>0 AND LEN(request("password"))>0 THEN
	
	'theURL=request.ServerVariables("HTTP_REFERER")
	'IF INSTR(request.ServerVariables("HTTP_REFERER"),"?")>0 THEN
	'	IF INSTR(request.ServerVariables("HTTP_REFERER"),"error")=0 THEN 
	'		theErrorURL=theURL & "&error=1"
	'	ELSE
	'		theErrorURL=theURL
	'	END IF
	'ELSE
	'	theErrorURL=theURL & "?error=1"
	'END IF
	'if request("username")="slettet" or request("password")="slettet" then
	'	theErrorURL=theURL & "?error=1"
	'	response.redirect theErrorURL
	'end if
	Session("username")=request("username")
	set rs=Server.createobject("ADODB.recordset")
	rs.activeconnection=MM_ecoinfo_STRING
	rs.Source="SELECT * FROM eisys_insiderusers WHERE username='" & request("username") & "' AND password='" & request("password") & "'"
	rs.Open()
	IF NOT rs.EOF THEN
		IF rs.Fields.Item("userlevel")=10 THEN
			session("eiuserid")=rs.Fields.Item("id").value
			response.cookies("eiuserid")=rs.Fields.Item("id").value
			response.cookies("eiorgname")="&Oslash;ko-net medarbejder"
			response.cookies("eiinsider")="1"
'			response.cookies("eiuserid").Expires=DateAdd("d",60,Date)
		ELSE
			response.cookies("eiorgid")=rs.Fields.Item("orgid").value
			orgid=rs.Fields.Item("orgid").value
			rs.close()
			rs.Source="SELECT name,firstname,lastname FROM eiorg_maindata WHERE id=" & orgid
			rs.open()
			IF LEN(rs.Fields.Item("name").value)=0 THEN
				response.cookies("eiorgname")=rs.Fields.Item("firstname").value & " " & rs.Fields.Item("lastname").value
			ELSE
				response.cookies("eiorgname")=rs.Fields.Item("name").value
			END IF
'			response.cookies("eiorgid").Expires=DateAdd("d",120,Date)
'			response.cookies("eiorgname").Expires=DateAdd("d",120,Date)
		END IF
		response.redirect "/home/index.asp"
	ELSE
		response.redirect theErrorURL
	END IF
ELSE
		response.redirect request.ServerVariables("HTTP_REFERER")
END IF
%>