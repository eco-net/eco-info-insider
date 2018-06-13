<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/connections/ecoinfo.asp"-->

<%
server.ScriptTimeout = 30000 ' st laaaaang tid af til upload (15 min.)


IF request("mail")=1 THEN
	IF CINT("0" & request("test"))=1 THEN
		TestNewsMail	
	ELSE
		SendNewsMail
	END IF
ELSEIF request("mail")=2 THEN
	IF CINT("0" & request("test"))=1 THEN
		TestInsiderMail	
	ELSE
		SendInsiderMail
	END IF
ELSEIF request("mail")=3 THEN
	IF CINT("0" & request("test"))=1 THEN
		TestFmpMail	
	ELSE
		SendFmpMail
	END IF
END IF

'response.redirect ("functions.asp")

SUB SendNewsMail
	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT email FROM eisys_newmail WHERE siteid=1"
	rs.open()
	Set msg = Server.CreateOBject( "JMail.Message" )
	msg.From = "eco-net@eco-net.dk"
	msg.Charset = "iso-8859-1"
	msg.Subject = request("headline")
	msg.Body = request("body")
	msg.FromName = "ko-info"
	WHILE NOT rs.EOF
		msg.recipients.clear()
		msg.AddRecipient rs.Fields.Item("email").value
		msg.Send("mail.eco-info.dk")
		rs.movenext
	WEND
	rs.close()
	rs=""
END SUB

SUB SendInsiderMail
	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT email,username,password FROM eisys_insiderusers WHERE email IS NOT NULL AND email<>''"

	IF InStr(request("body"),"#username#")>0 THEN do_u=1
	IF InStr(request("body"),"#password#")>0 THEN do_p=1
	rs.open()
	Set msg = Server.CreateOBject( "JMail.Message" )
	msg.From = "eco-net@eco-net.dk"
	msg.Charset = "iso-8859-1"
	msg.Subject = request("headline")
	msg.FromName = "ko-info"
	
	IF do_u OR do_p THEN
		theBody=request("body")
		WHILE NOT rs.EOF
			msg.Body=theBody
			IF do_u THEN msg.Body = replace(msg.Body,"#username#",rs.Fields.Item("username").value)
			IF do_p THEN msg.Body = replace(msg.Body,"#password#",rs.Fields.Item("password").value)
			msg.recipients.clear()
			msg.AddRecipient rs.Fields.Item("email").value
			on error resume next
			msg.Send("mail.eco-info.dk")
			if err then response.write "Kunne ikke sende til :" & rs("username") & ", " & rs("email") & "<br>"
			rs.movenext
		WEND
	ELSE
		msg.Body = request("body")
		WHILE NOT rs.EOF
			msg.recipients.clear()
			msg.AddRecipient rs.Fields.Item("email").value
			msg.Send("mail.eco-info.dk")
			on error resume next
			rs.movenext
		WEND
	END IF
	rs.close()
	rs=""
	response.write "<br>br><h3>Handlingen er fuldfrt.</h3>"
END SUB


SUB SendFmpMail

	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT mail FROM eisys_fmpmail"
	rs.open()
	Set msg = Server.CreateOBject( "JMail.Message" )
	msg.From = "eco-net@eco-net.dk"
	msg.Charset = "iso-8859-1"
	msg.Subject = request("headline")
	msg.Body = request("body")
	msg.FromName = "ko-net"
	WHILE NOT rs.EOF
		msg.recipients.clear()
		msg.AddRecipient rs.Fields.Item("mail").value
		on error resume next
		msg.Send("mail.eco-info.dk")
		rs.movenext
	WEND
	rs.close()
	rs=""
END SUB


SUB TestInsiderMail
	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT Count(*) as theCount FROM eisys_insiderusers WHERE email IS NOT NULL AND email<>''"
	rs.open()
	response.write "<h3> Test af Insidermail</h3>Mailen vil blive sendt til " & rs("theCount") & " personer.<br>" 
	response.write "Brug 'Vis kilde' for at se mailen i dens rette format, og husk ogs at tjekke den kopi der er sendt til eco-net@eco-net.dk.<br><br>Nedenfor vises de frste 50 mails.<br><br>"
	rs.close()

	rs.source="SELECT Top 50 email,username,password FROM eisys_insiderusers WHERE email IS NOT NULL AND email<>''"

	IF InStr(request("body"),"#username#")>0 THEN do_u=1
	IF InStr(request("body"),"#password#")>0 THEN do_p=1
	rs.open()
	
	IF do_u OR do_p THEN
		theBody=request("body")
		WHILE NOT rs.EOF
			theText=theBody
			IF do_u THEN theText = replace(theText,"#username#",rs.Fields.Item("username").value)
			IF do_p THEN theText = replace(theText,"#password#",rs.Fields.Item("password").value)
			response.write theText & chr(13) & "<br><br><br>" & chr(13) & chr(13) & chr(13)
			rs.movenext
		WEND
	ELSE
		theBody = request("body")
		WHILE NOT rs.EOF
			response.write theBody & chr(13) & "<br><br><br>" & chr(13) & chr(13) & chr(13)
			rs.movenext
		WEND
		theText=theBody
	END IF
	rs.close()	
	rs=""

	Set msg = Server.CreateOBject( "JMail.Message" )
	msg.From = "eco-net@eco-net.dk"
	msg.Charset = "iso-8859-1"
	msg.Subject = "TEST:: " & request("headline")
	msg.FromName = "ko-info"
	msg.Body = theText
	msg.AddRecipient "henrik@rediin.dk"
	msg.AddRecipient "eco-net@eco-net.dk"
	msg.Send("mail.eco-info.dk")
	set msg = nothing

	response.end
END SUB


SUB TestFmpMail

	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT Count(*) as theCount FROM eisys_fmpmail"

	response.write "<h3> Test af Filemaker Mail</h3>Mailen vil blive sendt til " & rs("theCount") & " personer.<br>" 
	response.write "Brug 'Vis kilde' for at se mailen i dens rette format, og husk ogs at tjekke den kopi der er sendt til eco-net@eco-net.dk.<br><br>Nedenfor vises den frste mail.<br><br>"
	rs.close()

	rs.source="SELECT TOP 1 mail FROM eisys_fmpmail"
	rs.open()
	WHILE NOT rs.EOF
		response.write request("body")& chr(13) & "<br><br><br>" & chr(13) & chr(13) & chr(13)
		rs.movenext
	WEND
	rs.close()
	rs=""

	Set msg = Server.CreateOBject( "JMail.Message" )
	msg.From = "eco-net@eco-net.dk"
	msg.Charset = "iso-8859-1"
	msg.Subject = "TEST:: " & request("headline")
	msg.Body = request("body")
	msg.FromName = "ko-net"
	msg.AddRecipient "eco-net@eco-net.dk"
	msg.Send("mail.eco-info.dk")
	set msg = nothing

	response.end
END SUB


SUB TestNewsMail

	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT Count(*) as theCount FROM eisys_newmail WHERE siteid=1"

	response.write "<h3> Test af News Mail</h3>Mailen vil blive sendt til " & rs("theCount") & " personer.<br>" 
	response.write "Brug 'Vis kilde' for at se mailen i dens rette format, og husk ogs at tjekke den kopi der er sendt til eco-net@eco-net.dk.<br><br>Nedenfor vises den frste mail.<br><br>"
	rs.close()

	rs.source="SELECT email FROM eisys_newmail WHERE siteid=1"
	rs.open()
	WHILE NOT rs.EOF
		response.write request("body")& chr(13) & "<br><br><br>" & chr(13) & chr(13) & chr(13)
		rs.movenext
	WEND
	rs.close()
	rs=""

	Set msg = Server.CreateOBject( "JMail.Message" )
	msg.From = "eco-net@eco-net.dk"
	msg.Charset = "iso-8859-1"
	msg.Subject = "TEST:: " & request("headline")
	msg.Body = request("body")
	msg.FromName = "ko-info"
	msg.AddRecipient "eco-net@eco-net.dk"
	msg.Send("mail.eco-info.dk")
	set msg = nothing

	response.end
END SUB

%>
