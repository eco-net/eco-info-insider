<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
' Åbnes med:
' window.open('test2.asp?fldname=t1','editor','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,width=500,height=250,top=200,left=100')"
'
' hvor fldname = navn på det tekstfelt teksten står i
' og fldnum = feltets index, hvis der er flere felter af samme navn.
%>
<html>
<head>
<title>Tekst formatering</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="/shared/styles.css">
</head>
<script language="JavaScript" type="text/JavaScript">
var t = opener.document.forms[0].<%=request("fldname")%><%if len(request("fldnum"))>0 then response.write "[" & request("fldnum") & "]"%>.value;

function doclose(){
opener.document.forms[0].<%=request("fldname")%><%if len(request("fldnum"))>0 then response.write "[" & request("fldnum") & "]"%>.value = document.frames("myEditor").document.frames("textEdit").document.body.innerHTML
window.close()
}
</script>

<body bgcolor="#EEEEEE" onload="document.frames('myEditor').document.frames('textEdit').document.body.innerHTML=t;">
<form name="editor" method="post" action="">
<table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
<tr>
<td class="functionsitem_inline" width="100%" height="100%">
<font color="#EEEEEE">
<!--#include virtual="/shared/editor/inc_textarea.htm"-->
</font>
</td></tr>
<tr><td height="25">
<div align="right">
<font color="#EEEEEE">
<input type="button" class="formbutton" onclick="window.close();" value="Fortryd">
<input type="button" class="formbutton" onclick="doclose();" value="OK">
</font></div></td>
</tr>
</table>
</form>
</body>
</html>
