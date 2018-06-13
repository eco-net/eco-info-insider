<script src="/log/optionfiles/selamtkomm.js"></script>
<script language="JavaScript">
function submitSearch(theform) {
	if(theform.cotime.selectedIndex>0)theform.cotimename.value=theform.cotime.options[theform.cotime.selectedIndex].text;
	if(theform.cocat.selectedIndex>0)theform.cocatname.value=theform.cocat.options[theform.cocat.selectedIndex].text;
	if(theform.cosbj.selectedIndex>0)theform.cosbjname.value=theform.cosbj.options[theform.cosbj.selectedIndex].text;
	if(theform.cokomm.selectedIndex>0)theform.cokommname.value=theform.cokomm.options[theform.cokomm.selectedIndex].text;
	if(theform.coamt.selectedIndex>0)theform.coamtname.value=theform.coamt.options[theform.coamt.selectedIndex].text;
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
<input type="hidden" name="cotimename" value="">
<input type="hidden" name="cocatname" value="">
<input type="hidden" name="cosbjname" value="">
<input type="hidden" name="cokommname" value="">
<input type="hidden" name="coamtname" value="">
<input type="hidden" name="sektion" value="kurser">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;VIS KURSER</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top" colspan="2"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td>
<span class="formLabeltext">I perioden</span>...<br>
<select name="cotime" class="formInputobject">
<script src="/log/optionfiles/oktime_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) i kategorien</span>...<br>
<select name="cocat" class="formInputobject">
<script src="/log/insider/ei/menufiles/coursecat_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) under emnet</span>...<br>
<select name="cosbj" class="formInputobject">
<script src="/log/insider/ei/menufiles/sbj_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) i regionen </span>...<br>
<select name="coamt" class="formInputobject" onChange="populateSel(this.form.cokomm,this.form.coamt.options[this.form.coamt.selectedIndex].value)">
<script src="/log/optionfiles/amter_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) i kommunen</span>...<br>
<select name="cokomm" class="formInputobject">
<script src="/log/optionfiles/kommuner_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) med teksten</span>...<br>
<input type="text" name="cotext" class="formInputobject">
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
