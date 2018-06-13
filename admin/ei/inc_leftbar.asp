
<div align="center"><br>
<b>Administration af <br>
&Oslash;ko-info</b></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td colspan="2"><img src="/shared/graphics/spacer.gif" width="3" height="40"></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td> 
<%
IF level1=-1 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader"><a href="/admin/ei/vindue/functions.asp">Vindue</a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td> 
<%
IF level1=0 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader"><a href="/admin/ei/tema/functions.asp">Tema </a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td height="18"> 
<%
IF level1=1 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader" height="18"><a href="/admin/ei/homepage/functions.asp">Forsiden</a></td>
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
<td class="sidebarHeader"><a href="/admin/ei/sections/functions.asp">Sektioner</a></td>
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
<td class="sidebarHeader"><a href="/admin/ei/cats/functions.asp">Kategorier og 
emneord</a></td>
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
<td class="sidebarHeader"><a href="/admin/ei/ads/functions.asp">Annoncer & Bannere</a></td>
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
<td class="sidebarHeader"><a href="/admin/ei/mails/functions.asp">Emails</a></td>
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
<td class="sidebarHeader"><a href="/admin/ei/help/functions.asp?intern=0">&Oslash;ko-net 
hj&aelig;lpetekst</a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr height="20"> 
<td> 
<%
IF level1=7 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader"><a href="/admin/ei/help/functions.asp?intern=1">Bruger 
hj&aelig;lpetekst</a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"></td>
</tr>
</table>
