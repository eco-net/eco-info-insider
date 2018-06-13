<!--#include virtual="/shared/inc_adm_access.asp" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="388" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td valign="top"> 
<%
SET fs = CreateObject("Scripting.FileSystemObject")
filename="/log/ei/home/vindue/inc_leader_" & request("id") & ".txt"
Set ts=fs.OpenTextFile(Server.mapPath(filename))
response.write ts.ReadAll()


%>

<table border="0" cellpadding="0" cellspacing="0" align="center">
<tr> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader"> &nbsp;&nbsp;SEKTIONER</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
<%
SET fs = CreateObject("Scripting.FileSystemObject")
filename="/log/ei/home/tema/inc_sections_" & request("id") & ".txt"
Set ts=fs.OpenTextFile(Server.mapPath(filename))
response.write ts.ReadAll()


%>
</td>
</tr>
<tr> 
<td valign="bottom" align="left"> 

</td>
</tr>
</table>


</body>
</html>
