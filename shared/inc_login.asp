
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<table width="180" border="0" cellspacing="0" cellpadding="0">
<tr>
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr>
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;LOG IND</td>
</tr>
<tr>
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top" colspan="2"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
<tr>
<td>
<%IF LEN(request.cookies("eiuserid"))=0 AND LEN(request.cookies("eiorgid"))=0 THEN%>
<form name="form1" method="post" action="/dgs/new.asp?newuser=1">
For at kunne bruge &Oslash;ko-info Insider og vores databaser, skal du tilmeldes De Gr&oslash;nne Sider.
<div align="center">
<input type="submit" name="Submit" value="Tilmeld" class="formbutton">
</div>
</form>
<%END IF%>
<form name="form1" method="post" action="/shared/login.asp" onsubmit="if(this.username.value.length==0 || this.password.value.length==0){alert('Begge felter skal udfyldes.');return false};else{return true;}">
<span class="SidebarHeader">Log ind</span><br>
<%IF LEN(request.cookies("eiuserid"))=0 AND LEN(request.cookies("eiorgid"))=0 THEN%>
Log ind for at oprette og &aelig;ndre arrangementer og publikationer.<br>
<br>
<%IF LEN(request("error"))>0 THEN%>
<font color="red">Der findes ingen bruger med de indtastede adgangsoplysninger !</font>
<br><br>
<%END IF%>
<span class="formLabeltext">
Brugernavn:</span>
<input type="text" name="username" class="formInputobject">
<span class="formLabeltext"><br>
Adgangskode:</span><br>
<input type="text" name="password" class="formInputobject">
<input type="hidden" name="referer" value="<%=request.ServerVariables("SCRIPT_NAME")%>">
<div align="center"><br>
<input type="submit" name="Button" value="Log ind" class="formSubmitbutton">
</div>
<%ELSE%>
Du er logget ind som <b><%=request.cookies("eiorgname")%></b>.<br>
<%IF request.cookies("eiinsider")="1" THEN%>
<br>
<input type="button" name="Submit" value="Log ind p&aring; &Oslash;ko-info" class="forminputobject" onClick="this.form.action='http://www.eco-info.dk/shared/eixxx.asp';this.form.target='_blank';this.form.submit();">
<%END IF%>
<%END IF%>
</form>
</td>
</tr>
<tr>
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr><td>
<%
a=1
if a=1 then %>
<%IF LEN(request.cookies("eiuserid"))=0 AND LEN(request.cookies("eiorgid"))=0 THEN%>
<form name="form2" action="forgotten_mail.asp" method="post" target="_blank">
<span class="SidebarHeader"><br>
Glemt brugerdata? </span> <br>
<br>
<input type="text" name="mailadr" value="din_mail@dresse" class="formInputobject">

<p align="center"> 
<input type="submit" name="Submit2" value="Sp&oslash;rg &Oslash;ko-net" class="formbuttonWide">
</p>
</form>
<%end if %>
<%end if %>
<table width="91%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td> 
<div align="center"> <br>
<input type="button" name="Button2" value="Logout" onClick="MM_goToURL('parent','/shared/logout.asp');return document.MM_returnValue" class="formSubmitbutton">
</div>
</td>
</tr>
</table>
</td></tr>
<tr>
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</table>
</td>
</tr>
</table>
