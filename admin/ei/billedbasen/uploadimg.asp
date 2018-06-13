<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp"-->
<%

DIM theImagePath, UploadRequest,filename
Set UploadRequest = CreateObject("Scripting.Dictionary")
DIM  RequestBin
RequestBin = Request.BinaryRead(Request.TotalBytes)
PureUploadSetup

theImagePath="/admin/ei/billedbasen/3col/"

filename=theImagePath & SaveFile_ASP(theImagePath,1)


response.write filename
%>