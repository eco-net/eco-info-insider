<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
response.write "Start"

	set rs=Server.createobject("ADODB.recordset")
	rs.activeconnection=MM_ecoinfo_STRING
	rs.Source="SELECT * FROM site_statistik"
	rs.Open()
	while not rs.EOF
	response.write MM_ecoinfo_STRING
	response.write rs("page") & "-" & rs("dato") & "-" & rs("site") & "<br>"
	rs.MoveNext
	wend
	rs.Close
	response.write "Slut"
%>