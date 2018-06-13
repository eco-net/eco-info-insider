<%

x=(timer mod #bannercount#) + 1

SET fs = CreateObject("Scripting.FileSystemObject")
filename=Server.MapPath("/log/ei/inc_banner_" & CStr(x) & ".txt")
Set ts=fs.OpenTextFile(filename)
response.write ts.ReadAll()
ts=""
fs=""

%>