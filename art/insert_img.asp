<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim rsPagedata__theID
rsPagedata__theID = "0"
if (request("id")  <> "") then rsPagedata__theID = request("id") 
%>
<%
set rsPagedata = Server.CreateObject("ADODB.Recordset")
rsPagedata.ActiveConnection = MM_ecoinfo_STRING
rsPagedata.Source = "SELECT title,author  FROM eiart_maindata  WHERE id=" + Replace(rsPagedata__theID, "'", "''") + ""
rsPagedata.CursorType = 0
rsPagedata.CursorLocation = 2
rsPagedata.LockType = 3
rsPagedata.Open()
rsPagedata_numRows = 0
%>
<html>
<head>
<title>Inds&aelig;t JPG billede</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
<script language="JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0

document.form1.img_size.value=jsStr;

document.form1.submit();
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="left" class="contentHeader2">Inds&aelig;t JPG-Billede i Artikel<br>
<%=(rsPagedata.Fields.Item("title").Value)%> </p>
<p align="left">Der kan inds&aelig;ttes <span class="listheader">JPG</span> billeder:<br>
efter teksten (max. 360 pixel bred),<br>
til h&oslash;jre for teksten (max. 160 pixel bred)</p>
<form name="form1" method="post" action="/admin/ei/billedbasen/upload_art_jpg.asp" enctype="multipart/form-data">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td height="358"> 
<p> Find dit billede<br>
<input type="file" name="file" class="formInputobjectLong">
</p>
<p> Skriv billedtekst<br>
<input type="text" name="subtext" class="formInputobject">
</p>
<p>Angiv kilde<br>
<% If Not rsPagedata.EOF Or Not rsPagedata.BOF Then %>
<input value="<%=(rsPagedata.Fields.Item("author").Value)%>" type="text" name="source" class="formInputobject">
<% End If ' end Not rsPagedata.EOF Or NOT rsPagedata.BOF %>
<input type="hidden" name="img_size">
<input type="hidden" name="artid" value="<%=request("id")%>">
</p>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td width="200" bgcolor="#FFFFCC"> 
<p class="listheader" align="center"><%=(rsPagedata.Fields.Item("title").Value)%></p>
<p align="center">her er din tekst</p>
<p align="center"><a href="javascript:MM_callJS('3')"><img src="/shared/graphics/layout/insert_img_big.jpg" width="150" height="150"></a></p>
</td>
<td width="100" bgcolor="#FFFF66"> 
<div align="center"><a href="javascript:MM_callJS('R')"><img src="/shared/graphics/layout/insert_img.jpg" width="75" height="75"></a></div>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
<p>&nbsp;</p>
</body>
</html>
<%
rsPagedata.Close()
%>
