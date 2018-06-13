
<link rel="stylesheet" href="/shared/styles.css" type="text/css">

<div align="center"><br>
<b>Billedbasen</b></div>
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
IF level1=1 THEN
	response.write "<img src=""/shared/graphics/layout/arrows_fwd.gif"" width=""10"" height=""7"">"
ELSE
	response.write "<img src=""/shared/graphics/spacer.gif"" width=""10"" height=""18"">"
END IF
%>
</td>
<td class="sidebarHeader"><a href="index.asp">Se Billeder</a></td>
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
<td class="sidebarHeader"><a href="insert.asp">Opret Nyt Billede</a><br>
GIF, BMP, <br>
Splash, Branding</td>
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
<td class="sidebarHeader"><a href="insert_jpg.asp">Opret JPG Billede</a><br>
Auto-kopiering </td>
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
<td class="sidebarHeader"><a href="cat.asp">Kategorier</a></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>

</table>

