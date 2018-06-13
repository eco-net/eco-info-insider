<html><!-- #BeginTemplate "/Templates/noheader.dwt" --><!-- DW6 -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Bekræftelse</title>
<!-- #EndEditable --> 
<script src="/shared/common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="7" marginwidth="0" marginheight="7">
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" name="Pagetable">
<tr> 
<td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
<td width="100%" valign="top"> 
<!-- START HEADER -->
<table width="100%" border="0" cellspacing="0" cellpadding="0" name="Header">
<tr> 
<td valign="top" rowspan="3" width="180" heigth="33"><img src="/shared/graphics/logo.gif" width="180" height="33"></td>
<td height="17">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td align="left"><img src="/shared/graphics/logo2.gif" width="21" height="16"></td>
</tr>
</table>
</td>
</tr>
<tr>
<td valign="top" width="99%" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr>
<td height="16"><br></td>
</tr>
</table>
<!-- END HEADER -->

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td width="100%" valign="top" name="maincontent">
<table border="0" cellspacing="0" cellpadding="10" width="100%" name="Contentarea">
<tr><td valign="top">
<!-- #BeginEditable "maincontent" -->
<table border="0" cellspacing="0" cellpadding="10" width="100%"><tr><td>
<%IF request("what")=1 THEN%>
Email-adressen <b><%=request("mail")%></b> er <b>tilmeldt</b> EcoInfo's nyhedsmail.
<%ELSE%>
Email-adressen <b><%=request("mail")%></b> er <b>ikke længere tilmeldt</b> EcoInfo's nyhedsmail.
<%END IF%>
</td></tr><tr><td align="right">
<a href="javascript:window.close();">Luk vindue</a>
</td></tr></table>
<!-- #EndEditable -->
</td></tr>
</table>
</td>
</tr>
</table>
</td>
<td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
</tr>
<tr height="1"> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="3"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
</body>
<!-- #EndTemplate --></html>
<!--#include virtual="/shared/log.asp"-->
