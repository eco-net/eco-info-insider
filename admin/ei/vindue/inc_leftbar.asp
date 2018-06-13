<!--#include virtual="/Connections/ecoinfo.asp" -->
<div align="center"><br>
<b>Vindue  p&aring;<br>
&Oslash;ko-info</b></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td  class="sidebarHeader" colspan="2"> 
<p>&nbsp;</p>
<p><a href="/admin/ei/vindue/functions.asp">V&aelig;lg Vindue </a> </p>
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
<td class="sidebarHeader"><a href="edit.asp?id=<%=request("id")%>&name=<%=request("name")%>">Rediger 
Leder p&aring; Vindue </a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td valign="top"> 
<%
IF level1=2 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader">Rediger Sektioner <br>
- <a href="edit_sec_dgs.asp">De 
Gr&oslash;nne sider</a><br>
- <a href="edit_sec_ok.asp">&Oslash;ko-Kalender</a><br>
- <a href="edit_sec_dgb.asp">Det 
Gr&oslash;nne Bibliotek</a><br>
- <a href="edit_sec_opsl.asp">Opslagstavle</a><br>
- <a href="edit_sec_news.asp">Nyheder</a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td> 
<%
IF level1=3 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader"><a href="edit_artikler.asp?id=<%=request("id")%>&name=<%=request("name")%>">Rediger 
Artikler p&aring; Vinduesside</a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td> 
<%
IF level1=5 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader"><a href="index_show.asp?id=<%=request("id")%>" target="_blank">Vis 
forside-eksempel med dette vindue </a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
