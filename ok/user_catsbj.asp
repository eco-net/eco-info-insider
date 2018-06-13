<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim rsCategories__ColParam
rsCategories__ColParam = "0"
if (request("id") <> "") then rsCategories__ColParam = request("id")
%>
<%
set rsCategories = Server.CreateObject("ADODB.Recordset")
rsCategories.ActiveConnection = MM_ecoinfo_STRING
rsCategories.Source = "SELECT c.id,c.name  FROM eical_cat_maindata c LEFT JOIN eical_r_cat o ON c.id=o.catid  WHERE o.calid=" + Replace(rsCategories__ColParam, "'", "''") + ""
rsCategories.CursorType = 0
rsCategories.CursorLocation = 2
rsCategories.LockType = 3
rsCategories.Open()
rsCategories_numRows = 0
%>
<%
Dim rsSubjects__ColParam
rsSubjects__ColParam = "0"
if (request("id") <> "") then rsSubjects__ColParam = request("id")
%>
<%
set rsSubjects = Server.CreateObject("ADODB.Recordset")
rsSubjects.ActiveConnection = MM_ecoinfo_STRING
rsSubjects.Source = "SELECT s.id,s.name  FROM eical_r_sbj o LEFT JOIN eisbj_maindata s ON o.sbjid=s.id  WHERE o.calid=" + Replace(rsSubjects__ColParam, "'", "''") + ""
rsSubjects.CursorType = 0
rsSubjects.CursorLocation = 2
rsSubjects.LockType = 3
rsSubjects.Open()
rsSubjects_numRows = 0
%>
<script src="user_multiselect.js"></script>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td> 
<form name="form1" method="post" action="user_catsbj_mail.asp">
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td> 
<p><span class="contentHeader1">Her kan du &oslash;nske nye kategorier og emneord 
til dit Arrangement.</span></p>
<p><span class="listheader">Arrangementet: <%=request("name")%></span> 
<input type="hidden" name="arrname" value="<%=request("name")%>">
<input type="hidden" name="id" value="<%=request("id")%>">
<br>
V&aelig;lg og send til &Oslash;ko-net, s&aring; retter vi.</p>
<div align="center"></div>
</td>
</tr>
</table>
<br>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td><span class="formLabeltext">Alle kategorier:</span><br>
<select name="allCat" class="formTextobjectLow" size="5" onDblClick="addValue(this.form.allCat,this.form.selCat);">
<script src="/log/insider/ei/menufiles/okcat_options.js"></script>
</select>
<br>
Dobbeltklik p&aring; en kategori for at v&aelig;lge den</td>
<td><span class="formLabeltext">Valgte kategorier:</span><br>
<select name="selCat" class="formTextobjectLow" size="5" onDblClick="removeValue(this.form.selCat);" multiple>
<%
While (NOT rsCategories.EOF)
%>
<option value="<%=(rsCategories.Fields.Item("name").Value)%>"><%=(rsCategories.Fields.Item("name").Value)%></option>
<%
  rsCategories.MoveNext()
Wend
If (rsCategories.CursorType > 0) Then
  rsCategories.MoveFirst
Else
  rsCategories.Requery
End If
%>
</select>
<br>
Dobbeltklik p&aring; en kategori for at fjerne den </td>
</tr>
</table>
<br>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td><span class="formLabeltext">Alle emneord:</span><br>
<select name="allSbj" class="formTextobjectLow" size="5"  onDblClick="addValue(this.form.allSbj,this.form.selSbj);">
<script src="/log/insider/ei/menufiles/sbj_options.js"></script>
</select>
<br>
Dobbeltklik p&aring; et emneord for at v&aelig;lge det</td>
<td><span class="formLabeltext">Valgte Emneord:</span><br>
<select name="selSbj" class="formTextobjectLow" size="5" onDblClick="removeValue(this.form.selSbj);" multiple>
<%
While (NOT rsSubjects.EOF)
%>
<option value="<%=(rsSubjects.Fields.Item("name").Value)%>"><%=(rsSubjects.Fields.Item("name").Value)%></option>
<%
  rsSubjects.MoveNext()
Wend
If (rsSubjects.CursorType > 0) Then
  rsSubjects.MoveFirst
Else
  rsSubjects.Requery
End If
%>
</select>
<br>
Dobbeltklik p&aring; et emneord for at fjerne det </td>
</tr>
</table>
<br>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td> 
<div align="center"> 
<input type="submit" name="Submit" value="Send" class="formSubmitbutton" onClick="selectselected(this.form.selCat);selectselected(this.form.selSbj)">
</div>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
</body>
</html>
<%
rsCategories.Close()
%>
<%
rsSubjects.Close()
%>
