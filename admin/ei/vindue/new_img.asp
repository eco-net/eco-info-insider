<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="/shared/styles.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
//alert(document.form1.radiobutton.checked);
if(document.form1.radiobutton[1].checked==1){
theURL=theURL + "?sizes=5"
}
else {theURL=theURL + "?sizes=4"}

  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body>
<form name="form1" method="post" action="act_image.asp">

<table width="300"  border="0" align="center">
  <tr>
    <td bgcolor="#CCCCCC"><div align="center">
      <p class="formLabeltext">I<span class="formLabeltext">nds&aelig;t Billede fra Billedbasen</span></p>
      <p>&nbsp;</p>
    </div></td>
  </tr>
  <tr bgcolor="#CCCCCC">
    <td><div align="center"><span class="formLabeltext">V&aelig;lg kolonne: </span><br>
      </div>
        <table width="190" border="1" align="center">
          <tr>
            <td width="25%"><div align="center">
                <input name="radiobutton" type="radio" value="1" checked>
            </div></td>
            <td width="50%"><div align="center">
                <input name="radiobutton" type="radio" value="2">
            </div></td>
            <td width="25%"><div align="center">
                <input name="radiobutton" type="radio" value="3">
            </div></td>
          </tr>
      </table>
    <br></td>
  </tr>
  <tr bgcolor="#CCCCCC">
    <td><div align="center">
        <p><a href="/admin/ei/billedbasen/index.asp?sizes=2" target="_blank"><span class="listheader"><br>
          </span></a><span class="listheader">Find Billede:</span> 
		  <img src="/shared/graphics/layout/img.gif" width="29" height="24" onClick="MM_openBrWindow('/admin/ei/billedbasen/index.asp','','')">		  <br>
		  Skriv Billednr:
          <input type="text" name="img_id" size="4">
</p>
        <p>
          <span class="formLabeltext">Billedtekst:</span>          
          <input name="img_text" type="text" id="img_text">
</p>
    </div></td>
  </tr>
  <tr bgcolor="#CCCCCC">
    <td><p align="center"><span class="formLabeltext">R&aelig;kkef&oslash;lge:</span><br>
          <input name="sort" type="text" value="1" size="2">
    </p>
      <p align="center">
        <input name="Button" type="button" class="formSubmitbutton" value="Opret" onClick="this.form.submit();">
        <input name="Button2" type="button" class="formSubmitbutton" value="Fortryd" onClick="javascript:history.go(-1)">
        <input name="name" type="hidden" id="name" value="<%=request("name")%>">
        <input name="page_id" type="hidden" id="page_id" value="<%=request("page_id")%>">
      </p></td>
  </tr>
</table>
</form>
</body>
</html>
