<script src="/log/optionfiles/selamtkomm.js"></script>
<script language="JavaScript">
function submitSearch(theform) {
	if(theform.artcat.selectedIndex>0)theform.artcatname.value=theform.artcat.options[theform.artcat.selectedIndex].text;
	if(theform.artsbj.selectedIndex>0)theform.artsbjname.value=theform.artsbj.options[theform.artsbj.selectedIndex].text;
	return true
}
</script>
<%IF LEN(searchDescr)>0 THEN%>
<table width="180" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;AKTUEL S&Oslash;GNING</td>
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
<%=searchDescr%>
</td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</table>
</td>
</tr>
</table>
<%END IF%>
<table width="180" border="0" cellspacing="0" cellpadding="0">
<form name="theform" method="post" action="list.asp" onSubmit="return submitSearch(document.theform);">
<input type="hidden" name="artcatname" value="">
<input type="hidden" name="artsbjname" value="">
<input type="hidden" name="sektion" value="art">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;VIS PUBLIKATIONER</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top" colspan="2"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td><span class="formLabeltext">I kategorien</span>...<br>
<select name="artcat" class="formInputobject">
<script src="/log/insider/ei/menufiles/artcat_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) under emnet</span>...<br>
<select name="artsbj" class="formInputobject">
<script src="/log/insider/ei/menufiles/sbj_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) med teksten</span>...<br>
<input type="text" name="arttext" class="formInputobject">
</td>
</tr>
<tr> 
<td align="center"> <br>
<input type="submit" value="S&oslash;g" class="formSubmitbutton">
</td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</table>
</td>
</tr>
</form>
</table>

