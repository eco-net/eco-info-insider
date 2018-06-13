<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/sqlcheck.asp"-->
<%
name=request("name")
page_id=request("page_id")
item_id=request("item_id")

%>
<%
set rsVin = Server.CreateObject("ADODB.Recordset")
rsVin.ActiveConnection = MM_ecoinfo_STRING
rsVin.CursorType = 0
rsVin.CursorLocation = 2
rsVin.LockType = 3
rsVin.Source = "SELECT * FROM eivindue_item WHERE id=" & item_id
rsVin.Open()
%>
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
                <input name="radiobutton" type="radio" value="1" <% if rsVin("col")=1 then response.write "checked" %>>
            </div></td>
            <td width="50%"><div align="center">
                <input name="radiobutton" type="radio" value="2" <% if rsVin("col")=2 then response.write "checked" %>>
            </div></td>
            <td width="25%"><div align="center">
                <input name="radiobutton" type="radio" value="3" <% if rsVin("col")=3 then response.write "checked" %>>
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
          <input name="img_id" type="text" value="<%=rsVin("img_id")%>" size="4" >
</p>
        <p>
          <span class="formLabeltext">Billedtekst:</span>          
          <input name="img_text" type="text" id="img_text" value="<%=rsVin("img_text")%>">
</p>
    </div></td>
  </tr>
  <tr bgcolor="#CCCCCC">
    <td><p align="center"><span class="formLabeltext">R&aelig;kkef&oslash;lge:</span><br>
          <input name="sort" type="text" value="<%=rsVin("thesort")%>" size="2">
    </p>
      <p align="center">
        <input name="Button" type="button" class="formSubmitbutton" value="Ret" onClick="this.form.submit();">
        <input name="Button2" type="button" class="formSubmitbutton" value="Fortryd" onClick="javascript:history.go(-1)">
        <input name="name" type="hidden" id="name" value="<%=request("name")%>">
        <input name="page_id" type="hidden" id="page_id" value="<%=request("page_id")%>">
        <input name="update" type="hidden" id="update" value="1">
        <input name="item_id" type="hidden" id="item_id" value="<%=request("item_id")%>">
      </p></td>
  </tr>
</table>
</form>
</body>
</html>
<% rsVin.Close %>