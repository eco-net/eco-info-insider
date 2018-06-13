<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/connections/ecoinfo.asp"-->

<%
server.ScriptTimeout = 30000 ' sæt laaaaang tid af til upload (15 min.)


	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT orgid FROM eisys_insiderusers WHERE email IS NOT NULL AND email<>'' AND nomail<1"

	rs.open()
			
		WHILE NOT rs.EOF
			response.write(rs("orgid") & "<br>")
			rs.movenext
		WEND
	
	rs.close()
	rs=""
	response.write "<br>br><h3>Handlingen er fuldført.</h3>"



%>
