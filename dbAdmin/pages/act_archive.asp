<!--#include virtual="/shared/sqlcheck.asp"-->
<%
DIM theCat, theType, theName, showLog, themessage, theComm, theTitle, theCode
IF len(request("thetext"))=0 THEN
	theCode="%"
ELSE
	theCode=request("thetext")
END IF
IF CInt(request("delid"))>1 THEN
'**** Delete the statement
	Set theComm = Server.CreateObject("ADODB.Command")
	theComm.ActiveConnection = MM_dbAdmin_STRING
	theComm.CommandText = "DELETE FROM sqlstatements WHERE id=" & CStr(request("delid"))
	theComm.Execute
	theComm.ActiveConnection.Close
END IF
IF len(request("thefilter"))=0 THEN
	theCat="x"
	theType="0"
	theName="x"
	showLog=1
	theTitle="Logged staments"
ELSEIF request("theFilter")="a" THEN
'**** Show all saved
	theCat="_"
	theType="0"
	IF theCode<>"%" THEN
		theName=theCode
	ELSE
		theName="__"
	END IF
	showLog=0
	theTitle="All saved"
ELSEIF request("theFilter")="x" THEN
'**** Show executables
	theCat="_"
	theType="1"
	IF theCode<>"%" THEN
		theName=theCode
	ELSE
		theName="__"
	END IF
	showLog=0
	theTitle="Executables"
ELSEIF request("theFilter")="t" THEN
'**** Show templates
	theCat="_"
	theType="2"
	IF theCode<>"%" THEN
		theName=theCode
	ELSE
		theName="__"
	END IF
	showLog=0
	theTitle="Templates"
ELSEIF request("theFilter")="i" THEN
'**** Show infopages
	theCat="_"
	theType="3"
	IF theCode<>"%" THEN
		theName=theCode
	ELSE
		theName="__"
	END IF
	showLog=0
	theTitle="Info-pages"
ELSEIF request("theFilter")="l" THEN
'**** Show log
	theCat="x"
	theType="0"
	IF theCode<>"%" THEN
		theName=theCode
	ELSE
		theName="x"
	END IF
	showLog=1
	theTitle="Logged statements"
ELSEIF request("theFilter")="xl" THEN
'**** Empty log
	theCat="x"
	theType="0"
	theName="x"
	showLog=1
	theTitle="Clear log"
	Set theComm = Server.CreateObject("ADODB.Command")
	theComm.ActiveConnection = MM_dbAdmin_STRING
	theComm.CommandText = "DELETE FROM sqlstatements WHERE name='x'"
	theComm.Execute
	theComm.ActiveConnection.Close
	themessage="The log has been succesfully deleted"
ELSE
	theCat=request("theFilter")
	theType="0"
	IF theCode<>"%" THEN
		theName=theCode
	ELSE
		theName="x"
	END IF
	showLog=0
	theTitle="Saved statements in the category '" & theCat & "'"
END IF
'response.write theCat & " , " & theType & " , " & theName & " , " & CStr(showLog)
%>