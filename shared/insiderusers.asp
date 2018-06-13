<!--include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->

<%
Function regUser(username,orgid,page,sek,app_id,app_title)
'loginusername,new=1/edit=2/del=3,dgs/dgb/ok/kurser/udd/art,kort_id
DIM theComm,dato
dato=CStr(FormatDateTime(Date,2))
IF LEN(request.cookies("eiuserid"))=0 THEN
    Set theComm = Server.CreateObject("ADODB.Command")
    theComm.ActiveConnection = MM_ecoinfo_STRING
    theComm.CommandText = "INSERT INTO eisys_insideruser_stat (username,orgid,page,sek,app_id,dag,app_title,dato)  VALUES ('" &_
		username & "'," & orgid & ",'" & page & "','" & sek & "'," & app_id & "," & datepart("d",Date) & ",'" & app_title & "','" & dato & "')"
	
	theComm.Execute
    theComm.ActiveConnection.Close
end if
End Function
%>
