<script src="/log/optionfiles/selamtkomm.js"></script>
<script language="JavaScript">
function submitSearch(theform) {
	if(theform.dgscat.selectedIndex>0)theform.dgscatname.value=theform.dgscat.options[theform.dgscat.selectedIndex].text;
	if(theform.dgssbj.selectedIndex>0)theform.dgssbjname.value=theform.dgssbj.options[theform.dgssbj.selectedIndex].text;
	if(theform.dgskomm.selectedIndex>0)theform.dgskommname.value=theform.dgskomm.options[theform.dgskomm.selectedIndex].text;
	if(theform.dgsamt.selectedIndex>0)theform.dgsamtname.value=theform.dgsamt.options[theform.dgsamt.selectedIndex].text;
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
<input type="hidden" name="dgscatname" value="">
<input type="hidden" name="dgssbjname" value="">
<input type="hidden" name="dgskommname" value="">
<input type="hidden" name="dgsamtname" value="">
<input type="hidden" name="sektion" value="dgs">
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;VIS ORGANISATIONER</td>
</tr>
<tr> 
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top" colspan="2"> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td><span class="formLabeltext">I kategorien</span>...<br>
<select name="dgscat" class="formInputobject">
<script src="/log/insider/ei/menufiles/dgscat_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) under emnet</span>...<br>
<select name="dgssbj" class="formInputobject">
<script src="/log/insider/ei/menufiles/sbj_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) i regionen </span>...<br>
<select name="dgsamt" class="formInputobject" onChange="populateSel(this.form.dgskomm,this.form.dgsamt.options[this.form.dgsamt.selectedIndex].value)">
<script src="/log/optionfiles/amter_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) i kommunen</span>...<br>
<select name="dgskomm" class="formInputobject">
<script src="/log/optionfiles/kommuner_options.js"></script>
</select>
<br>
<span class="formLabeltext">(og) med teksten</span>...<br>
<input type="text" name="dgstext" class="formInputobject">
<br>
<input name="submit" type="submit" class="formSubmitbutton" value="S&oslash;g">
<br>
<hr>
<input name="stopped" type="checkbox" id="stopped" value="1">
<strong>Find oph&oslash;rte</strong><br>
<br>
<span class="formLabeltext">Søg på ID-numre i Filemaker<br>(indtast [adskil med linieskifte] eller Sæt ind fra fmp)</span>...<br>
<textarea name="dgsid" class="formInputobject" rows="2" wrap="VIRTUAL"></textarea>
<span class="formLabeltext">Søg på ID-numre på nettet<br>(indtast [adskil med linieskifte])</span>...<br>
<textarea name="netid" class="formInputobject" rows="2" wrap="VIRTUAL"></textarea>
<br>
NB ! Andre felter bruges ikke ved søgning på oph&oslash;rte eller ID.</td>
</tr>
<tr> 
<td align="center"> <br>
<input type="submit" value="S&oslash;g" class="formSubmitbutton">&nbsp;
<input type="submit" value="Find ID'er" class="formSubmitbutton" onClick="this.form.action=this.form.action+'?getids=1'">
<br>
<br>
<input type="button" name="Button" value="Ny organisation" class="formInputobject" onClick="window.location='new.asp?'+'<%=MM_KeepBoth%>'">
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
