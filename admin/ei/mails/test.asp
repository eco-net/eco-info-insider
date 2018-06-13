<%
Style="<style type='text/css'>BODY, TABLE, TD {font-family: Verdana, Arial, Helvetica, sans-serif;font-size:11px;}</style>"
gAfsenderNavn="Øko-net"
theTo=Array("peter@pyf.dk","web@eco-net.dk","eco-net@eco-net.dk","lmn@eco-net.dk","fejl127@mail.dk")
gServer="mx.ngoweb.dk"
thesender="nyhedsmail@eco-net.dk"
BODY="<table width='100%' border='0' cellpadding='5'><tr><td><strong>Nyhedsmail fra &Oslash;ko-net </strong></td><td><img src='http://www.eco-net.dk/shared/graphics/ECO-NET-logo.gif' alt='&Oslash;ko-net' width='180' height='58' /></td>  </tr><tr><td><p>Nu skal I h&oslash;re</p><p>bla. bla.</p><p>Hilsen &Oslash;ko-net   </p></td><td>&nbsp;</td>  </tr></table>"
sendErrorCount=0
sendCount=0
Set msg = Server.CreateOBject( "CDO.Message" )
	msg.From = gAfsenderNavn & "<" & thesender & ">"
	msg.Subject = "Testmail"
	msg.HTMLBody = "<html><head>" & Style & "</head><body>" & BODY & "</body></html>"
	For i=0 to 4
		email=theTo(i)
		if len(email)>5 then
			'msg.recipients.clear()
			'msg.AddRecipient email
			msg.To=email
			on error resume next
			'msg.Send(gServer)
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
		
	Next
	
%>
