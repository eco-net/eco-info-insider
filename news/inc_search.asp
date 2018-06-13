<script src="/log/optionfiles/selamtkomm.js"></script>
<script language="JavaScript">
function submitSearch(theform) {
	if(theform.dgbcat.selectedIndex>0)theform.dgbcatname.value=theform.dgbcat.options[theform.dgbcat.selectedIndex].text;
	if(theform.dgbsbj.selectedIndex>0)theform.dgbsbjname.value=theform.dgbsbj.options[theform.dgbsbj.selectedIndex].text;
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
<input type="hidden" name="dgbcatname" value="">
<input type="hidden" name="dgbsbjname" value="">
<input type="hidden" name="sektion" value="dgb">
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
<select name="dgbcat" class="formInputobject">
<script src="/log/optionfiles/dgbcat_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) under emnet</span>...<br>
<select name="dgbsbj" class="formInputobject">
<script src="/log/optionfiles/sbj_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) med teksten</span>...<br>
<input type="text" name="dgbtext" class="formInputobject">
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

