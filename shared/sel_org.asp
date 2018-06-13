<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->

<%
Dim rsPageData__ColParam
rsPageData__ColParam = "%"
if (request("searchval") <> "") then rsPageData__ColParam = request("searchval")
%>

<%
IF LEN(request("searchval"))>0 THEN
	set rsPageData = Server.CreateObject("ADODB.Recordset")
	rsPageData.ActiveConnection = MM_ecoinfo_STRING
	rsPageData.Source = "SELECT id,name, firstname, lastname FROM eiorg_maindata WHERE name LIKE '%" + Replace(rsPageData__ColParam, "'", "''") + "%' OR firstname LIKE '%" + Replace(rsPageData__ColParam, "'", "''") + "%' OR lastname LIKE '%" + Replace(rsPageData__ColParam, "'", "''") + "%' ORDER by name, lastname"
	rsPageData.CursorType = 0
	rsPageData.CursorLocation = 2
	rsPageData.LockType = 3
	rsPageData.Open()
	rsPageData_numRows = 0
END IF
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsPageData_numRows = rsPageData_numRows + Repeat1__numRows
%>

<html><!-- #BeginTemplate "/Templates/noheader.dwt" --><!-- DW6 -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Ecoinfo</title>
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
<form name="form1" method="post" action="">
Vis organisationer der matcher: 
<br>
<input type="text" name="searchval">
<input type="submit" name="Submit" value="S&oslash;g ..." class="formbutton">
<% 
IF LEN(request("searchval"))>0 THEN
If Not rsPageData.EOF Or Not rsPageData.BOF Then %>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td><b>Fundne organisationer</b></td>
</tr>

<% 
While ((Repeat1__numRows <> 0) AND (NOT rsPageData.EOF)) 
%>
<tr <%IF Repeat1__index MOD 2 = 0 THEN response.write "bgcolor=""#CCCCCC"""%>>
<td height="18">
<%
IF LEN(rsPageData("name"))>2 THEN
	thename=rsPageData("name")
ELSE
	thename=rsPageData("firstname") & " " & rsPageData("lastname")
END IF
response.write thename
%>
</td>
<td align="right" height="18">
&nbsp;&nbsp;<input type="button" value="V&aelig;lg" class="formbutton" onClick="opener.setdgs(<%=rsPageData("id") & ",'" & replace(replace(thename,"'","\'"),chr(13) & chr(10),"") & "'"%>);">
</td>
</tr>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsPageData.MoveNext()
Wend
%>
</table>
<% End If ' end Not rsPageData.EOF Or NOT rsPageData.BOF 
END IF
%>
</form>
<br>
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
<%
IF LEN(request("searchval"))>0 THEN
rsPageData.Close()
END IF
%>

<!--#include virtual="/shared/log.asp"-->

