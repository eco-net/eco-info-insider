<script src="/log/optionfiles/selamtkomm.js"></script>
<script language="JavaScript">
function submitSearch(theform) {
	
	if(theform.udcat.selectedIndex>0)theform.udcatname.value=theform.udcat.options[theform.udcat.selectedIndex].text;
	if(theform.udsbj.selectedIndex>0)theform.udsbjname.value=theform.udsbj.options[theform.udsbj.selectedIndex].text;
	
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

<input type="hidden" name="udcatname" value="">
<input type="hidden" name="udsbjname" value="">
<input type="hidden" name="sektion" value="udd">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;VIS UDDANNELSER</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top" colspan="2"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td> <span class="formLabeltext"> i kategorien</span>...<br>
<select name="udcat" class="formInputobject">
<script src="/log/insider/ei/menufiles/uddcat_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) under emnet</span>...<br>
<select name="udsbj" class="formInputobject">
<script src="/log/insider/ei/menufiles/sbj_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) med teksten</span>...<br>
<input type="text" name="udtext" class="formInputobject">
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
