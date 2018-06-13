<!--#include virtual="/shared/sqlcheck.asp"-->
<%
DIM theComm, isRs, theSQL, objType, rsData

IF LEN(request("del"))>0 THEN

	set rs=server.createobject("ADODB.command")
	rs.activeconnection=MM_ecoinfo_STRING
	IF request("delField")="id" THEN
		rs.commandtext="Delete FROM " & request("tablename") & " WHERE id IN (" & request("del") & ")"
	ELSE
		response.write request("del") & "<br>"
		theDel="'" & replace(request("del"),", ", "','") & "'"
		response.write theDel & "<br>"
		rs.commandtext="Delete FROM " & request("tablename") & " WHERE " & request("delField") & " IN (" & theDel & ")"
	END IF
	rs.execute
	rs.activeconnection.close()
	set rs=nothing
END IF

set rsData = Server.CreateObject("ADODB.Recordset")
rsData.ActiveConnection= MM_ecoinfo_STRING
'rsData.ActiveConnection = MM_voresDebat_STRING
rsData.Source = "SELECT * FROM " & request("tablename")
rsData.Open()

'**** Reading data into an array
theData=rsData.GetRows()
theheader="<tr class=""listlabelbg"">"
theheader=theheader & "<td class=""formlabelsleft"">Delete</td>"
i=0
For each item in rsData.Fields
	IF i=0 THEN theDelField=item.name
	theheader=theheader & "<td class=""formlabelsleft"">" & item.name & "</td>"
	i=1
NEXT		
theheader=theheader & "</tr>"

rsData.Close()
%>