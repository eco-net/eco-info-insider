<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/connections/ecoinfo.asp"-->

<%
response.buffer = false
server.ScriptTimeout = 3000000 ' st laaaaang tid af til upload (15 min.)

DIM email, sendCount, sendErrorCount, gServer
sendCount=0
sendErrorCount=0
gServer="mx.ngoweb.dk"
gAfsenderNavn = request("afsendernavn")

IF request("mail")=1 THEN
	IF CINT("0" & request("test"))=1 THEN
		TestNewsMail request("sender")
	ELSE
		SendFmpMail request("sender")
		SendNewsMail request("sender")
	END IF
ELSEIF request("mail")=2 THEN
	IF CINT("0" & request("test"))=1 THEN
		TestInsiderMail request("sender")
	ELSE
		SendInsiderMail request("sender")
	END IF
ELSEIF request("mail")=3 THEN
	IF CINT("0" & request("test"))=1 THEN
		TestFmpMail request("sender")	
	ELSE
		SendFmpMail request("sender")
	END IF
END IF

'response.redirect ("functions.asp")

SUB SendNewsMail(thesender)
	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT email FROM eisys_newmail WHERE siteid=1"
	rs.open()
	Set msg = Server.CreateObject( "CDO.Message" )
	msg.From = gAfsenderNavn & "<" & thesender & ">"
	msg.Subject = request("headline")
	msg.Body = replace(request("body"),chr(13)&chr(11),chr(11))
	msg.Body = replace(msg.Body,chr(11)&chr(13),chr(11))
	msg.Body = replace(msg.Body,chr(13),"")
	WHILE NOT rs.EOF
		email=rs("email")
		if len(email)>5 then
			'msg.recipients.clear()
			'msg.AddRecipient email
			msg.To=email
			on error resume next
			msg.Send
			if err then 
				sendErrorCount=sendErrorCount+1
				response.write sendErrorCount&"  <strong>Kunne ikke sende til :"&email&"</strong><br>"
			else
				sendCount=sendCount+1
				response.write sendCount&"  Sendt til: "&email&"<br>"
			end if				
		else
			response.write "<b>IKKE SENDT TIL: " & email & ". ADRESSEN ER IKKE KORREKT</b><br>"
		end if
		rs.movenext
	WEND
	rs.close()
	rs=""
	response.write "<h3>Rapport: Tilmeldte på Øko-info.</h3>Mailen er sendt til: "&sendCount&" adresser.<br>Mailen kunne ikke sendes til: "&sendErrorCount&" adresser."
END SUB

SUB SendInsiderMail(thesender)
	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT i.email,i.username,i.password,o.name FROM eisys_insiderusers i inner join eiorg_mindata o WHERE i.email IS NOT NULL AND i.email<>'' AND i.nomail<1"

	IF InStr(request("body"),"#username#")>0 THEN do_u=1
	IF InStr(request("body"),"#password#")>0 THEN do_p=1
	IF InStr(request("body"),"#organisation#")>0 THEN do_o=1
	
	rs.open()
	Set msg = Server.CreateOBject( "CDO.Message" )
	msg.From = gAfsenderNavn & "<" & thesender & ">"
	msg.Subject = request("headline")
	msg.Body = replace(request("body"),chr(13)&chr(11),chr(11))
	msg.Body = replace(msg.Body,chr(11)&chr(13),chr(11))
	msg.Body = replace(msg.Body,chr(13),"")
	
	IF do_u OR do_p OR do_o THEN
		WHILE NOT rs.EOF
			email=rs("email")
			if len(email)>5 then
				msg.Body=theBody
				IF do_u THEN msg.Body = replace(msg.Body,"#username#",rs.Fields.Item("username").value)
				IF do_p THEN msg.Body = replace(msg.Body,"#password#",rs.Fields.Item("password").value)
				IF do_o THEN msg.Body = replace(msg.Body,"#organisation#",rs.Fields.Item("name").value)
				'msg.recipients.clear()
				'msg.AddRecipient email
				msg.To=email
				on error resume next
				msg.Send
				if err then 
					sendErrorCount=sendErrorCount+1
					response.write sendErrorCount&"  <strong>Kunne ikke sende til :"&email&"</strong><br>"
				else
					sendCount=sendCount+1
					response.write sendCount&"  Sendt til: "&email&"<br>"
				end if				
			end if
			rs.movenext
		WEND
	ELSE
		WHILE NOT rs.EOF
			email=rs("email")
			if len(email)>5 then
				'msg.recipients.clear()
				'msg.AddRecipient email
				msg.To=email
				on error resume next
				msg.Send
				if err then 
					sendErrorCount=sendErrorCount+1
					response.write sendErrorCount&"  <strong>Kunne ikke sende til :"&email&"</strong><br>"
				else
					sendCount=sendCount+1
					response.write sendCount&"  Sendt til: "&email&"<br>"
				end if				
			else
				response.write "<b>IKKE SENDT TIL: " & email & ". ADRESSEN ER IKKE KORREKT</b><br>"
			end if
			rs.movenext
		WEND
	END IF
	rs.close()
	rs=""
	response.write "<h3>Rapport: Insiderbrugere.</h3>Mailen er sendt til: "&sendCount&" adresser.<br>Mailen kunne ikke sendes til: "&sendErrorCount&" adresser."
END SUB


SUB SendFmpMail(thesender)
	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT mail FROM eisys_fmpmail"
	rs.open()
	Set msg = Server.CreateOBject( "CDO.Message" )
	msg.From = gAfsenderNavn & "<" & thesender & ">"
	msg.Subject = request("headline")
	msg.Body = replace(request("body"),chr(13)&chr(11),chr(11))
	msg.Body = replace(msg.Body,chr(11)&chr(13),chr(11))
	msg.Body = replace(msg.Body,chr(13),"")
	x=0
	WHILE NOT rs.EOF
		email=rs("mail")
		if len(email)>5 then
			'msg.recipients.clear()
			'msg.AddRecipient email
			msg.To=email
			on error resume next
			msg.Send
			if err then 
				sendErrorCount=sendErrorCount+1
				response.write sendErrorCount&"  <strong>Kunne ikke sende til :"&email&"</strong><br>"
			else
				sendCount=sendCount+1
				response.write sendCount&"  Sendt til: "&email&"<br>"
			end if
		else
			response.write "<b>IKKE SENDT TIL: " & email & ". ADRESSEN ER IKKE KORREKT</b><br>"
		end if
		rs.movenext
	WEND
	rs.close()
	rs=""
	response.write "<h3>Rapport: Adresser fra Filemaker.</h3>Mailen er sendt til: "&sendCount&" adresser.<br>Mailen kunne ikke sendes til: "&sendErrorCount&" adresser."
END SUB


SUB TestInsiderMail(thesender)
	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT Count(*) as theCount FROM eisys_insiderusers WHERE email IS NOT NULL AND email<>''"
	rs.open()
	response.write "<h3> Test af Insidermail</h3>Mailen vil blive sendt til " & rs("theCount") & " personer.<br>" 
	response.write "Mailen er sendt til eco-net@eco-net.dk.<br><br>"
	rs.close()

	rs.source="SELECT TOP 1 i.email,i.username,i.password,o.name FROM eisys_insiderusers i inner join eiorg_maindata o ON i.orgid=o.id WHERE i.email IS NOT NULL AND i.email<>'' AND i.nomail<1"

	IF InStr(request("body"),"#username#")>0 THEN do_u=1
	IF InStr(request("body"),"#password#")>0 THEN do_p=1
	IF InStr(request("body"),"#organisation#")>0 THEN do_o=1
'	response.write "<bR>"&rs.source
'	response.end
	rs.open()
	
	IF do_u OR do_p OR do_o THEN
	thetext=replace(request("body"),chr(13),chr(11))
		IF do_u THEN theText = replace(theText,"#username#",rs.Fields.Item("username").value)
		IF do_p THEN theText = replace(theText,"#password#",rs.Fields.Item("password").value)
		IF do_o THEN theText = replace(theText,"#organisation#",rs.Fields.Item("name").value)
	ELSE
		theText = replace(request("body"),chr(13),chr(11))
	END IF
	rs.close()	
	rs=""

	Set msg = Server.CreateOBject( "CDO.Message" )
	msg.From = gAfsenderNavn & "<" & thesender & ">"
	msg.Subject = request("headline")
	msg.Body = theText
	msg.Body = replace(msg.Body,chr(13)&chr(11),chr(11))
	msg.Body = replace(msg.Body,chr(11)&chr(13),chr(11))
	msg.Body = replace(msg.Body,chr(13),"")
	'msg.AddRecipient "henrik@rediin.dk"
	'msg.AddRecipient "eco-net@eco-net.dk"
	msg.To="eco-net@eco-net.dk"
	msg.Send
	set msg = nothing

	response.end
END SUB


SUB TestFmpMail(thesender)

	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT Count(*) as theCount FROM eisys_fmpmail"
	rs.open()
	thecount=rs("theCount")
	rs.close()
	rs=""

	Set msg = Server.CreateOBject( "CDO.Message" )
	msg.From = gAfsenderNavn & "<" & thesender & ">"
	msg.Subject = request("headline")
	msg.Subject = "TEST FMPMail:: " & request("headline")
	msg.Body = replace(request("body"),chr(13)&chr(11),chr(11))
	msg.Body = replace(msg.Body,chr(11)&chr(13),chr(11))
	msg.Body = replace(msg.Body,chr(13),"")
	'msg.AddRecipient "eco-net@eco-net.dk"
	'msg.AddRecipient "henrik@rediin.dk"
	'msg.AddRecipient "hemure@gmail.com"
	msg.To="eco-net@eco-net.dk"
	msg.Send
	'msg.Send("mail.net-server.dk")
	set msg = nothing

	response.write "<h3> Test af Filemaker Mail</h3>Mailen vil blive sendt til " & theCount & " personer.<br>" 
	response.write "Mailen er sendt til eco-net@eco-net.dk.<br><br><a href=""javascript:history.go(-1);"">Tilbage</a>"

	response.end
END SUB


SUB TestNewsMail(thesender)

	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT Count(*) as theCount FROM eisys_newmail WHERE siteid=1"
	rs.open()
	tCount=rs("theCount")
	rs.close()
	rs.source="SELECT Count(*) as theCount FROM eisys_fmpmail"
	rs.open()
	tCount=rs("theCount")+tCount
	rs.close()
	response.write "<h3> Test af News Mail</h3>Mailen vil blive sendt til " & tCount & " personer.<br>" 
	response.write "Mailen er sendt til eco-net@eco-net.dk.<br><br><a href=""javascript:history.go(-1);"">Tilbage</a>"
	rs=""

	Set msg = Server.CreateOBject( "CDO.Message" )
	msg.From = gAfsenderNavn & "<" & thesender & ">"
	msg.Subject = request("headline")
	msg.Subject = "TEST:: " & request("headline")
	msg.Body = replace(request("body"),chr(13)&chr(11),chr(11))
	msg.Body = replace(msg.Body,chr(11)&chr(13),chr(11))
	msg.Body = replace(msg.Body,chr(13),"")
	'msg.AddRecipient "web@eco-net.dk"
	'msg.AddRecipient "eco-net@eco-net.dk"
	msg.To="eco-net@eco-net.dk"
	msg.Send
	set msg = nothing

	response.end
END SUB

%>
