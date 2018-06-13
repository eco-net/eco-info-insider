<%
response.write chr(13) & "<!-- START Tabs -->" & chr(13)
response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" bgcolor=""#FFFFFF"">"
response.write "<tr valign=""middle"">"

response.write chr(13) & "<!-- 1. tab -->" & chr(13)
if curTab=1 then
	response.write "<td><img src=""/dbadmin/images/layout/tableft_act.gif"" border=""0"" height=""20"" width=""5""></td>"
	response.write "<td background=""/dbadmin/images/layout/tab_act.gif"" class=""menu1active"" nowrap>SQL</td>"
	response.write "<td><img src=""/dbadmin/images/layout/tabright_act.gif"" border=""0"" height=""20"" width=""5""></td>"
else
	response.write "<td><img src=""/dbadmin/images/layout/tableft.gif"" border=""0"" height=""20"" width=""5""></td>"
	response.write "<td background=""/dbadmin/images/layout/tab.gif"" class=""menu1"" nowrap><a href=""sql.asp"">SQL</a></td>"
	response.write "<td><img src=""/dbadmin/images/layout/tabright.gif"" border=""0"" height=""20"" width=""5""></td>"
end if

response.write chr(13) & "<!-- 2. tab -->" & chr(13)
if curTab=2 then
	response.write "<td><img src=""/dbadmin/images/layout/tableft_act.gif"" border=""0"" height=""20"" width=""5""></td>"
	response.write "<td background=""/dbadmin/images/layout/tab_act.gif"" class=""menu1active"" nowrap>Result</td>"
	response.write "<td><img src=""/dbadmin/images/layout/tabright_act.gif"" border=""0"" height=""20"" width=""5""></td>"
else
	response.write "<td><img src=""/dbadmin/images/layout/tableft.gif"" border=""0"" height=""20"" width=""5""></td>"
	response.write "<td background=""/dbadmin/images/layout/tab.gif"" class=""menu1"" nowrap><a href=""result.asp"">Result</a></td>"
	response.write "<td><img src=""/dbadmin/images/layout/tabright.gif"" border=""0"" height=""20"" width=""5""></td>"
end if

response.write chr(13) & "<!-- 3. tab -->" & chr(13)
if curTab=3 then
	response.write "<td><img src=""/dbadmin/images/layout/tableft_act.gif"" border=""0"" height=""20"" width=""5""></td>"
	response.write "<td background=""/dbadmin/images/layout/tab_act.gif"" class=""menu1active"" nowrap>Archive</td>"
	response.write "<td><img src=""/dbadmin/images/layout/tabright_act.gif"" border=""0"" height=""20"" width=""5""></td>"
else
	response.write "<td><img src=""/dbadmin/images/layout/tableft.gif"" border=""0"" height=""20"" width=""5""></td>"
	response.write "<td background=""/dbadmin/images/layout/tab.gif"" class=""menu1"" nowrap><a href=""archive.asp"">Archive</a></td>"
	response.write "<td><img src=""/dbadmin/images/layout/tabright.gif"" border=""0"" height=""20"" width=""5""></td>"
end if

response.write chr(13) & "<!-- 4. tab -->" & chr(13)
if curTab=4 then
	response.write "<td><img src=""/dbadmin/images/layout/tableft_act.gif"" border=""0"" height=""20"" width=""5""></td>"
	response.write "<td background=""/dbadmin/images/layout/tab_act.gif"" class=""menu1active"" nowrap>Browse DB Structure</td>"
	response.write "<td><img src=""/dbadmin/images/layout/tabright_act.gif"" border=""0"" height=""20"" width=""5""></td>"
else
	response.write "<td><img src=""/dbadmin/images/layout/tableft.gif"" border=""0"" height=""20"" width=""5""></td>"
	response.write "<td background=""/dbadmin/images/layout/tab.gif"" class=""menu1"" nowrap><a href=""struct_tables.asp"">Browse DB Structure</a></td>"
	response.write "<td><img src=""/dbadmin/images/layout/tabright.gif"" border=""0"" height=""20"" width=""5""></td>"
end if

response.write chr(13) & "<!-- remaing space -->" & chr(13)
response.write "<td width=""90%"" background=""/dbadmin/images/layout/menuspace.gif""><img src=""/dbadmin/images/layout/spacer.gif"" border=""0"" width=""1"" height=""1""></td>"
response.write "<td><img src=""/dbadmin/images/layout/menuspaceright.gif"" border=""0"" width=""2"" height=""20""></td>"
response.write "</tr>"
response.write "</table>"
response.write chr(13) & "<!-- END Tabs -->" & chr(13)
%>

<%
IF curtab=4 THEN
	response.write chr(13) & "<!-- START Submenu -->" & chr(13)
	response.write "<table border=""0"" width=""750"" cellpadding=""0"" cellspacing=""0"">"
	response.write "<tr valign=""middle"">"
	response.write chr(13) & "<!-- left border -->" & chr(13)
	response.write "<td align=""left""><img src=""/dbadmin/images/layout/smleft.gif"" border=""0"" height=""20"" width=""5""></td>"
	
	response.write chr(13) & "<!-- menues -->" & chr(13)
	if curTab=4 then
		if curSub=1 then
			response.write "<td class=""menu2active"" nowrap background=""/dbadmin/images/layout/sm.gif"">Tables & Views"
		else
			response.write "<td class=""menu2"" nowrap background=""/dbadmin/images/layout/sm.gif"">"
			response.write "<a href=""struct_tables.asp"">Tables & Views</a>"
		end if
		
		response.write "&nbsp;&nbsp;" & chr(13)
		response.write "</td>"
		
		if curSub=2 then
			response.write "<td class=""menu2active"" nowrap background=""/dbadmin/images/layout/sm.gif"">Procedures"
		else
			response.write "<td class=""menu2"" nowrap background=""/dbadmin/images/layout/sm.gif"">"
			response.write "<a href=""struct_procedures.asp"">Procedures</a>"
		end if
		
		response.write "&nbsp;&nbsp;" & chr(13)
		response.write "</td>"
	END IF

	response.write "<td align=""right"" width=""90%"" background=""/dbadmin/images/layout/sm.gif"">"
	response.write "<img src=""/dbadmin/images/spacer.gif"" height=""1"" width=""1"">"
	response.write "</td>"
	response.write chr(13) & "<!-- right border -->" & chr(13)
	response.write "<td align=""right"" class=""menu2""><img src=""/dbadmin/images/layout/sm_other_r.gif"" border=""0"" height=""20"" width=""5""></td>"
	response.write "</tr>"
	response.write chr(13) & "<!-- END Submenu -->" & chr(13)
	response.write "</table>"
END IF
%>



<table border="0" width="750" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr>
<td background="/dbadmin/images/layout/page_l.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
<td width="99%"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="1" height="3"></td>
<td background="/dbadmin/images/layout/page_r.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
</tr>
</table>