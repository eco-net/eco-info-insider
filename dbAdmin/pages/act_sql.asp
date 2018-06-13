<!--#include virtual="/shared/sqlcheck.asp"-->
<%
IF len(request("dosave"))>0 THEN
'**** the current record must be saved
	DIM theComm,saveas, theName, theCat, objType, themessage
	IF Len(request("theName"))=0 THEN theName="x" ELSE theName=request("theName")
	IF Len(request("Category"))=1 THEN theCat="x" ELSE theCat=request("Category")
	saveas=request("saveas")+0
	objType=CInt(request("objType"))
	
	Set theComm = Server.CreateObject("ADODB.Command")
	theComm.ActiveConnection = MM_ecoinfo_STRING
	'theComm.ActiveConnection = MM_dbAdmin_STRING
	theComm.CommandText = "UPDATE sqlstatements SET name='" & theName + "',code='" &_
	replace(request("theSql"),"'","''") & "',isLink=" & CStr(saveas) & ",objecttype=" &_
	CStr(objType) & ",category='" & theCat & "',description='" &_
	replace(request("thedescr"),"'","''") & "' " &_
	"WHERE id=" & request("id")	
	theComm.Execute
	theComm.ActiveConnection.Close

	IF len(request("exe"))=1 THEN
	'**** execute the code
		response.redirect("result.asp?id=" & Cstr(request("id")))
	ELSE
		themessage="The data has been saved."
	END IF
END IF
%>