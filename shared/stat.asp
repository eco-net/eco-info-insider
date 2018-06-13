
<!--#include virtual="/Connections/ecoinfo.asp" -->

<%

SetLocale(1030)
Dim referrer,remote,page,site,dato

FUNCTION FormatDateDB(theDate)
	IF LEN(theDate)>0 THEN
		theDate=CDate(theDate)
		FormatDateDB="{ts '" & datepart("yyyy",theDate) & "-" & right("0" & CStr(datepart("m",theDate)),2) & "-" &_
			right("0" & CSTR(datepart("d",theDate)),2) & " 00:00:00'}"
	ELSE
		FormatDateDB="''"
	END IF
END FUNCTION

Function register()
Session("reg")=1

referrer=Request.ServerVariables("HTTP_REFERER")
remote=Request.ServerVariables("REMOTE_ADDR")
page=Request.ServerVariables("SCRIPT_NAME")
dato=FormatDateDB(Now)
site="insider"
if remote<>"80.160.41.182" then
'writeDB()
docount()
end if

End Function

Function writeDB()

DIM theComm
    Set theComm = Server.CreateObject("ADODB.Command")
    theComm.ActiveConnection = MM_ecoinfo_STRING
    theComm.CommandText = "INSERT INTO site_statistik (referrer,remoteadr,page,dato,site)  VALUES ('" &_
		referrer & "','" & remote & "','" & page & "'," & dato & ",'" & site & "')"
	theComm.Execute
    theComm.ActiveConnection.Close
	
End Function
Function docount()
dim c
set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_ecoinfo_STRING
rs.Source = "SELECT *  FROM stat_counter  ORDER BY date DESC"
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 2
rs.Open()

if rs("date")<> date then
	rs.AddNew
	rs("date")=date
	rs("insider") = 1
	rs("bu")=0
	rs("ecoinfo")=0
	rs("econet")=0
	rs.Update
else
	if rs("insider")=0 then
		c=1
	else
		c=rs("insider") + 1
	end if
	rs("insider")=c
	rs.Update
end if
rs.Close()
end Function

if len(Session("reg"))=0 then

register()
end if
%>

