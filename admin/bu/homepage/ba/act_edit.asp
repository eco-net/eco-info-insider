<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="lib.asp" -->

<%
'**** Saving imagefile
DIM theImagePath, UploadRequest, filename
Set UploadRequest = CreateObject("Scripting.Dictionary")
DIM  RequestBin
RequestBin = Request.BinaryRead(Request.TotalBytes)
PureUploadSetup

filename=""
IF LEN(UploadRequest.Item("file").Item("Value"))>0 THEN 
	theImagePath="/log/bu/home/branding/"
	filename=theImagePath & SaveFile_ASP(theImagePath,1)
END IF
DoSave filename,2
%>