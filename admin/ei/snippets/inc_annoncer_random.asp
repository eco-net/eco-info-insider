<%
IF len(recinfo)>0 THEN
	response.write "<table width='181' height='100%' border='0' cellspacing='0' cellpadding='0'>" &_
		"<tr><td width='1' height='100%' background='/shared/graphics/layout/dots_vert.gif'><img src='/shared/graphics/spacer.gif' width='1' height='1'></td><td width='180' valign='top'>" &_
		"<table width='180' border='0' cellspacing='0' cellpadding='0'>" &_
		"<tr bgcolor='#FAFAF4'><td>" &_
		"<img src='/shared/graphics/layout/header_arrows.gif' width='22' height='22'></td>" &_
		"<td width='98%' class='sidebarHeader'>" &_
		"&nbsp;&nbsp;#Title#</td></tr><tr>" &_
		"<td colspan='2' height='1' background='/shared/graphics/layout/dots_horz.gif'>" &_
		"<img src='/shared/graphics/spacer.gif' width='3' height='1'></td></tr>" &_
		"<tr><td valign='top' colspan='2'>"
	'**** Writing contentheader
	response.write "<table width='90%' border='0' cellspacing='0' cellpadding='0' align='center'>" &_
		"<tr><td><img src='/shared/graphics/spacer.gif' width='3' height='8'></td></tr>"
ELSE
	response.write "<table width='180' border='0' cellspacing='0' cellpadding='0'>" &_
		"<tr><td colspan='2' height='1' background='/shared/graphics/layout/dots_horz.gif'><img src='/shared/graphics/spacer.gif' width='3' height='1'></td></tr>" &_	
		"<tr bgcolor='#FAFAF4'><td><img src='/shared/graphics/layout/header_arrows.gif' width='22' height='22'></td>" &_
		"<td width='98%' class='sidebarHeader'>&nbsp;&nbsp;#Title#</td></tr>" &_
		"<tr><td colspan='2' height='1' background='/shared/graphics/layout/dots_horz.gif'><img src='/shared/graphics/spacer.gif' width='3' height='1'></td></tr>" &_
		"<tr><td valign='top' colspan='2'><table width='90%' border='0' cellspacing='0' cellpadding='0' align='center'>" &_
		"<tr><td><img src='/shared/graphics/spacer.gif' width='3' height='8'></td></tr>"

END IF


x=(timer mod #addcount#) + 1

SET fs = CreateObject("Scripting.FileSystemObject")
filename=Server.MapPath("/log/ei/#sektion#/inc_add_" & CStr(x) & ".asp")
Set ts=fs.OpenTextFile(filename)
execute ts.ReadAll()
ts=""
x=(x mod #addcount#) + 1
filename=Server.MapPath("/log/ei/#sektion#/inc_add_" & CStr(x) & ".asp")
Set ts=fs.OpenTextFile(filename)
execute ts.ReadAll()
ts=""
fs=""

IF len(recinfo)>0 THEN
	'**** Content footer
	response.write "<tr><td><img src='/shared/graphics/spacer.gif' width='3' height='8'></td>" &_
		"</tr></table>"
END IF
'**** Writing Footer
response.write "</td></tr></table></td></tr></table>"
%>
