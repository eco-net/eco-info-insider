<!--#include virtual="/shared/sqlcheck.asp"-->
<%
DIM theComm, isRs, theSQL, objType, rsData
theSQL=request("thesql")
objType=Cint(request("objType"))

IF LEN(request("id"))>0 THEN
'**** Execute a stored statement
		set rsData = Server.CreateObject("ADODB.Recordset")
		rsData.ActiveConnection = MM_ecoinfo_STRING
		'rsData.ActiveConnection = MM_dbAdmin_STRING
		rsData.Source = "SELECT code, objectType FROM sqlstatements WHERE id=" _
			& Cstr(request("id"))
	'	response.write rsData.Source + "<br>"
		rsData.Open()
	
		'**** Reading data into an array
		theSQL=rsData.Fields.Item("code").value
		objType=rsData.Fields.Item("objectType").value
		rsData.Close()

ELSEIF LEN(theSQL)>0 THEN
'**** the entered code is saved/logged
	DIM saveas, theName, theCat
	IF Len(request("theName"))=0 THEN theName="x" ELSE theName=request("theName")
	IF Len(request("Category"))=1 THEN theCat="x" ELSE theCat=request("Category")
	
	saveas=request("saveas")+0
	objType=CInt(request("objType"))
    Set theComm = Server.CreateObject("ADODB.Command")
    theComm.ActiveConnection = MM_dbAdmin_STRING
    theComm.CommandText = "INSERT INTO sqlstatements (name, code, isLink, objecttype,category,description) VALUES ('" + _
		theName + "','" + replace(request("theSql"),"'","''") + "'," + _
		CStr(saveas) + "," + CStr(objType) + ",'" + theCat + "','" + _
		replace(request("thedescr"),"'","''") + "')"
'	response.write theComm.CommandText + "<br>"
'	theComm.Execute
    theComm.ActiveConnection.Close
	IF saveas>=2 THEN theSQL=""
END IF

IF len(thesql)>0 THEN
	'**** Executing the sqlcode
	IF objType=0 THEN
	'**** Create a recordset
		DIM theData
		isRs=1
		set rsData = Server.CreateObject("ADODB.Recordset")
		rsData.ActiveConnection = MM_ecoinfo_STRING
		'rsData.ActiveConnection = MM_voresDebat_STRING
		rsData.Source = thesql
	'	response.write rsData.Source + "<br>"
		rsData.Open()
	
		'**** Reading data into an array
		theData=rsData.GetRows()
		theheader="<tr class=""listlabelbg"">"
		FOr each item in rsData.Fields
			theheader=theheader & "<td class=""formlabelsleft"">" & item.name & "</td>"
		NEXT		
		theheader=theheader & "</tr>"
		rsData.Close()
	ELSE
	'**** Create a Command
		DIM theResponse
		isRs=0
		Set theComm = Server.CreateObject("ADODB.Command")
		theComm.ActiveConnection = MM_ecoinfo_STRING
		'theComm.ActiveConnection = MM_voresDebat_STRING
		theComm.CommandText = thesql
	'	response.write theComm.CommandText + "<br>"
		theComm.Execute
		theComm.ActiveConnection.Close
		theResponse="The statement has executed succesfully !"
	END IF
END IF
%>