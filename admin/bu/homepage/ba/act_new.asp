<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="lib.asp"-->

<%
'**** Saving imagefile
DIM theImagePath, UploadRequest,filename
Set UploadRequest = CreateObject("Scripting.Dictionary")
DIM  RequestBin
RequestBin = Request.BinaryRead(Request.TotalBytes)
PureUploadSetup

theImagePath="/log/bu/home/branding/"
filename=theImagePath & SaveFile_ASP(theImagePath,1)

DoSave filename,1
%>
