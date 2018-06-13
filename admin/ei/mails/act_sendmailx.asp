<!--include virtual="/shared/inc_adm_access.asp" -->
<!--include virtual="/connections/ecoinfo.asp"-->

<%
server.ScriptTimeout = 300000 ' st laaaaang tid af til upload (15 min.)

SUB SendInsiderMail()
	
	
	Set msg = Server.CreateOBject( "JMail.Message" )
	msg.From = "insider@eco-net.dk"
	msg.Charset = "iso-8859-1"
	msg.Subject = "headline"
	msg.FromName = "Øko-info"
	body="hej med dig her er en mail"
	frameld=Chr(11) & "<a href='http://www.eco-info.dk'>Frameld denne emailservice</a>"
	
		theBody=replace(body,chr(13),chr(11))
		theBody=theBody & frameld
			msg.Body=theBody
			'msg.recipients.clear()
			msg.AddRecipient "jj@eco-net.dk"
			
			msg.Send("mail.eco-info.dk")
			
			
		response.write msg.From & "<br>" & theBody & "<br>"
	
	response.write "<br><br><h3>Handlingen er fuldført.</h3>"
END SUB

SendInsiderMail
%>
