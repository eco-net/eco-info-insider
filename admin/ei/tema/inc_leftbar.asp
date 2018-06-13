<!--#include virtual="/Connections/ecoinfo.asp" -->
<div align="center"><br>
<b>Tema p&aring;<br>
&Oslash;ko-info</b></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td  class="sidebarHeader" colspan="2"> 
<p>&nbsp;</p>
<p><a href="/admin/ei/tema/functions.asp">V&aelig;lg tema</a> </p>
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
Leder p&aring; Forside</a></td>
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
- <a href="/admin/ei/tema/edit_sec_dgs.asp?id=<%=request("id")%>&name=<%=request("name")%>">De 
Gr&oslash;nne sider</a><br>
- <a href="/admin/ei/tema/edit_sec_ok.asp?id=<%=request("id")%>&name=<%=request("name")%>">&Oslash;ko-Kalender</a><br>
- <a href="/admin/ei/tema/edit_sec_dgb.asp?id=<%=request("id")%>&name=<%=request("name")%>">Det 
Gr&oslash;nne Bibliotek</a><br>
- <a href="/admin/ei/tema/edit_sec_opsl.asp?id=<%=request("id")%>&name=<%=request("name")%>">Opslagstavle</a><br>
- <a href="/admin/ei/tema/edit_sec_news.asp?id=<%=request("id")%>&name=<%=request("name")%>">Nyheder</a></td>
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
<td class="sidebarHeader"><a href="/admin/ei/tema/edit_artikler.asp?id=<%=request("id")%>&name=<%=request("name")%>">Rediger 
Artikler p&aring; Temaside</a></td>
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
<td class="sidebarHeader"><a href="/admin/ei/ads/functions.asp">Inds&aelig;t Annoncer 
p&aring; Temaside</a> </td>
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
forside-eksempel med dette tema</a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td> 
<%
IF level1=6 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader"><a href="/admin/ei/tema/afstemning.asp">Afstemning</a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
