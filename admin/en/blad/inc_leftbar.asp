<!--#include virtual="/Connections/ecoinfo.asp" -->

<div align="center"><br>
<b>&Oslash;ko-net<br>
Nyhedsblad</b></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td  class="sidebarHeader" colspan="2"> 
<p>&nbsp;</p>
<p><a href="/admin/en/index.asp">V&aelig;lg blad</a> </p>
</td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td> 
<%
IF level1=1 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader"><a href="edit.asp?id=<%=request("id")%>&name=<%=request("name")%>">Forside</a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td> 
<%
IF level1=2 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader"><a href="/admin/en/blad/kapitel.asp?id=<%=request("id")%>&name=<%=request("name")%>">Afsnit</a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td valign="top"> 
<%
IF level1=3 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader"><a href="/admin/en/blad/artikel.asp?id=<%=request("id")%>&name=<%=request("name")%>">Artikel</a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td> 
<%
IF level1=4 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader"><a href="pdf.asp?id=<%=request("id")%>&name=<%=request("name")%>">Pdf</a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td>&nbsp; </td>
<td class="sidebarHeader">Kolofon<br>
rettes i HTML</td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
