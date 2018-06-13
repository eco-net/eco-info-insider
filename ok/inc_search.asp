<script src="/log/optionfiles/selamtkomm.js"></script>
<script language="JavaScript">
function submitSearch(theform) {
	if(theform.oktime.selectedIndex>0)theform.oktimename.value=theform.oktime.options[theform.oktime.selectedIndex].text;
	if(theform.okcat.selectedIndex>0)theform.okcatname.value=theform.okcat.options[theform.okcat.selectedIndex].text;
	if(theform.oksbj.selectedIndex>0)theform.oksbjname.value=theform.oksbj.options[theform.oksbj.selectedIndex].text;
	if(theform.okkomm.selectedIndex>0)theform.okkommname.value=theform.okkomm.options[theform.okkomm.selectedIndex].text;
	if(theform.okamt.selectedIndex>0)theform.okamtname.value=theform.okamt.options[theform.okamt.selectedIndex].text;
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
<input type="hidden" name="oktimename" value="">
<input type="hidden" name="okcatname" value="">
<input type="hidden" name="oksbjname" value="">
<input type="hidden" name="okkommname" value="">
<input type="hidden" name="okamtname" value="">
<input type="hidden" name="sektion" value="ok">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;VIS ARRANGEMENTER</td>
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
<select name="oktime" class="formInputobject">
<script src="/log/optionfiles/oktime_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) i kategorien</span>...<br>
<select name="okcat" class="formInputobject">
<script src="/log/insider/ei/menufiles/okcat_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) under emnet</span>...<br>
<select name="oksbj" class="formInputobject">
<script src="/log/insider/ei/menufiles/sbj_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) i regionen </span>...<br>
<select name="okamt" class="formInputobject" onChange="populateSel(this.form.okkomm,this.form.okamt.options[this.form.okamt.selectedIndex].value)">
<script src="/log/optionfiles/amter_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) i kommunen</span>...<br>
<select name="okkomm" class="formInputobject">
<script src="/log/optionfiles/kommuner_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) med teksten</span>...<br>
<input type="text" name="oktext" class="formInputobject">
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
