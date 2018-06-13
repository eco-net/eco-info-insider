<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/common.asp"-->



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
<%
Dim rsImgR__theID
rsImgR__theID = "0"
if (request("id")  <> "") then rsImgR__theID = request("id") 
%>
<%
set rsImgR = Server.CreateObject("ADODB.Recordset")
rsImgR.ActiveConnection = MM_ecoinfo_STRING
rsImgR.Source = "SELECT *  FROM images i LEFT JOIN eiart_r_img r ON i.id=r.imgid  WHERE r.artid=" + Replace(rsImgR__theID, "'", "''") + " AND r.imgsize='R'  ORDER BY sort"
rsImgR.CursorType = 0
rsImgR.CursorLocation = 2
rsImgR.LockType = 3
rsImgR.Open()
rsImgR_numRows = 0
%>
<%
Dim rsImg3__theID
rsImg3__theID = "0"
if (request("id")   <> "") then rsImg3__theID = request("id")  
%>
<%
set rsImg3 = Server.CreateObject("ADODB.Recordset")
rsImg3.ActiveConnection = MM_ecoinfo_STRING
rsImg3.Source = "SELECT *  FROM images i LEFT JOIN eiart_r_img r ON i.id=r.imgid  WHERE r.artid=" + Replace(rsImg3__theID, "'", "''") + " AND r.imgsize='3'  ORDER BY sort"
rsImg3.CursorType = 0
rsImg3.CursorLocation = 2
rsImg3.LockType = 3
rsImg3.Open()
rsImg3_numRows = 0
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
rsImgR_numRows = rsImgR_numRows + Repeat2__numRows
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsImg3_numRows = rsImg3_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Rediger JPG billede</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
<script language="JavaScript">
<!--
function MM_callJS(subtext,img,sort) { //v2.0

	document.form1.update.value="1";
	document.form1.subtekst.value=subtext;
	alert(document.form1.subtekst.value)
	document.form1.sort.value=sort;
	document.form1.imgid.value=img;
	document.form1.submit();

}

function delimg(theform){
	theform.dodel.value='1';
	theform.submit();
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="left" class="contentHeader2">Rediger JPG-Billeder i Artikel<br>
<%=(rsPagedata.Fields.Item("title").Value)%> </p>
<p align="left">&nbsp;</p>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<th>Billeder under tekst</th>
<th>Billeder til højre for tekst</th>
</tr>
<tr> 
<td height="358" valign="top" bgcolor="#FFFFCC"> 
<% If Not rsImg3.EOF Or Not rsImg3.BOF Then %>
<br>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsImg3.EOF)) 
%>
<form method="post" action="saveImg.asp">
<div align="center"> <img src="/admin/ei/billedbasen/3col/<%=(rsImg3.Fields.Item("filename").Value)%> "> 
</div>
<p align="center"> Billedtekst <br>
R&aelig;kkef&oslash;lge 
<input value="<%=(rsImg3.Fields.Item("sort").Value)%>" type="text" name="sort" size="2">
<input value="<%=(rsImg3.Fields.Item("subtext").Value)%>" type="text" name="subtext" class="formInputobject">
<br>
<input type="submit" name="Submit3" value="Opdater" class="formSubmitbutton">
<input type="button" name="Del3" value="Slet" class="formSubmitbutton" onClick="delimg(this.form)">
<input type="hidden" name="id" value="<%=request("id")%>">
<input type="hidden" name="imgid" value="<%=rsImg3("imgid")%>">
<input type="hidden" name="dodel" value="0">
</p>
</form>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsImg3.MoveNext()
Wend
%>
<% End If ' end Not rsImg3.EOF Or NOT rsImg3.BOF %>
<input type="hidden" name="update" value="1">
</td>
<td height="358" valign="top" bgcolor="#FFFF99"> 
<div align="center"> 
<% If Not rsImgR.EOF Or Not rsImgR.BOF Then %>
<br>
<% 
While ((Repeat2__numRows <> 0) AND (NOT rsImgR.EOF)) 
%>
<form method="post" action="saveImg.asp">
<div align="center"> <img src="/admin/ei/billedbasen/rightcol/<%=(rsImgR.Fields.Item("filename").Value)%> "> 
</div>
<p align="center"> Billedtekst 
<input value="<%=(rsImgR.Fields.Item("subtext").Value)%>" type="text" name="subtext" class="formInputobject">
<br>
R&aelig;kkef&oslash;lge 
<input value="<%=(rsImgR.Fields.Item("sort").Value)%>" type="text" name="sort" size="2">
<br>
<input type="submit" name="SubmitR" value="Opdater" class="formSubmitbutton">
<input type="button" name="DelR" value="Slet" class="formSubmitbutton" onClick="delimg(this.form)">
<input type="hidden" name="id" value="<%=request("id")%>">
<input type="hidden" name="imgid" value="<%=rsImgR("imgid")%>">
<input type="hidden" name="dodel" value="0">
</p>
</form>
<% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsImgR.MoveNext()
Wend
%>
<% End If ' end Not rsImgR.EOF Or NOT rsImgR.BOF %>
<br>
</div>
</td>
</tr>
</table>

<p>&nbsp;</p>
</body>
</html>
<%
rsPagedata.Close()
%>
<%
rsImgR.Close()
%>
<%
rsImg3.Close()
%>
