<%
x=DatePart("s",date)
y=#count#
z=(x mod y)+1
filename=Server.MapPath("inc_branding_" & CStr(z) & ".txt")
SET fs = CreateObject("Scripting.FileSystemObject")
Set ts=fs.OpenTextFile(filename)
response.write ts.ReadAll()
ts=""
fs=""
%>