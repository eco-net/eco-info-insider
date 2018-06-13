<!--#include virtual="/shared/sqlcheck.asp"-->
<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Ecoinfo</title>
<script src="/shared/multiselect.js"></script>
<script src="lib.js"></script>
<script src="/shared/common.js"></script>
<script src="/shared/showeditor.js"></script>
<!-- #EndEditable --> 
<script src="/shared/common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script></head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="7" marginwidth="0" marginheight="7">
<table width="752" border="0" cellspacing="0" cellpadding="0" align="center" name="Pagetable">
<tr> 
<td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
<td width="750" valign="top"> 
<!-- START HEADER -->
<!-- #BeginLibraryItem "/Library/header.lbi" -->
<table width="750" border="0" cellspacing="0" cellpadding="0" name="Header">
<tr> 
<td valign="top" rowspan="3"><a href="/index.asp"><img src="/shared/graphics/logo.gif" width="180" height="33" border="0"></a> 
<table width="180" border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td width="8"><br>
</td>
<td class="sitetag"> Registrering af og med gr&oslash;nne gr&aelig;sr&oslash;dder, NGO'er, foreninger, firmaer m.fl.<br>
</td>
<td width="8"><br>
</td>
</tr>
</table>
</td>
<!--
<td valign="top" width="1"><br>
</td>
-->
<td height="17" align="right" colspan="4">

<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td align="left"><img src="/shared/graphics/logo2.gif" width="21" height="16"></td>
<td align="right"> <a href="/home/index.asp">Forside</a> | <a href="http://www.eco-info.dk" target="_blank">&Oslash;ko-info</a> 
| <a href="/home/about_1.asp">Om Øko-info Insider</a> | <a href="/home/searching.asp">Kom 
i gang</a> | <a href="#" onClick="window.open('/shared/help/help.asp','InsiderHelp','scrollbars=yes,resizable=yes,width=700,height=300');">Hj&aelig;lp</a></td>
</tr>
</table>
</td>
</tr>
<tr> 
<td valign="top" rowspan="2" width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="3"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td width="388" height="64"> 
<div align="center">
</div>
</td>
<td background="/shared/graphics/layout/dots_vert.gif" width="1"><br>
</td>
<td width="180" align="center" valign="top" background="/shared/graphics/layout/search_bkgrd.gif"> 
<table width="150" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
<tr> 
<td>
</td>
</tr>
</table>
</td>
</tr>

</table>
<!-- #EndLibraryItem --><!-- END HEADER -->
<!-- #BeginEditable "menu" -->
<%DIM curtab
curtab=1
%>
<!-- #BeginLibraryItem "/Library/menu.lbi" --><table width="750" border="0" cellspacing="0" cellpadding="0" name="Menu">
<tr>
<%
IF LEN(request.cookies("eiuserid"))=0 THEN 'Organisation eller ikke logget ind
	'-- tab1 --
	response.write "<td><img src=""/shared/graphics/menu/stretch.gif"" width=""40"" height=""22"">"
	IF curtab=1 THEN
		response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
			"<a href=""/dgs/index.asp""><img src=""/shared/graphics/menu/dgsu_on.gif"" width=""64"" height=""22"" border=""0""></a>"
	ELSE
		response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">" &_
			"<a href=""/dgs/index.asp""><img src=""/shared/graphics/menu/dgsu_off.gif"" width=""64"" height=""22"" border=""0""></a>"
	END IF
	'-- tab2 --
	IF curtab=2 THEN
		response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
			"<a href=""/ok/list.asp""><img src=""/shared/graphics/menu/ok_on.gif"" width=""100"" height=""22"" border=""0""></a>"
	ELSE
		IF curtab=1 THEN
			response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
		ELSE
			response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
		END IF
		response.write "<a href=""/ok/list.asp""><img src=""/shared/graphics/menu/ok_off.gif"" width=""100"" height=""22"" border=""0""></a>"
	END IF
	'-- tab3 --
	IF curtab=3 THEN
		response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
			"<a href=""/dgb/list.asp""><img src=""/shared/graphics/menu/dgb_on.gif"" width=""88"" height=""22"" border=""0""></a>"
	ELSE
		IF curtab=2 THEN
			response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
		ELSE
			response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
		END IF
		response.write "<a href=""/dgb/list.asp""><img src=""/shared/graphics/menu/dgb_off.gif"" width=""88"" height=""22"" border=""0""></a>"
	END IF
	'-- tab 4 --
	
	IF curtab=4 THEN
		response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
			"<a href=""/kurser/list.asp""><img src=""/shared/graphics/menu/kurser_on.gif"" width=""60"" height=""22"" border=""0""></a>"
	ELSE
		IF curtab=3 THEN
			response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
		ELSE
			response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
		END IF
		response.write "<a href=""/kurser/list.asp""><img src=""/shared/graphics/menu/kurser_off.gif"" width=""60"" height=""22"" border=""0""></a>"
	END IF
	'-- tab 7 --
	
	IF curtab=7 THEN
		response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
			"<a href=""/art/list.asp""><img src=""/shared/graphics/menu/artikler_on.gif"" width=""60"" height=""22"" border=""0""></a>" &_
			"<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
	
	ELSE
		IF curtab=4 THEN
			response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
		ELSE
			response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
		END IF
		response.write "<a href=""/art/list.asp""><img src=""/shared/graphics/menu/artikler_off.gif"" width=""60"" height=""22"" border=""0""></a>" 
		response.write "<img src=""/shared/graphics/menu/stretch.gif"" width=""30"" height=""22"">" 
	END IF
	response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""30"" height=""22"">" 
	response.write "<img src=""/shared/graphics/menu/stretch.gif"" width=""133"" height=""22""></td>"
	
ELSE 
'Øko-net sektretariatet
	'-- tab1 --
	response.write "<td><img src=""/shared/graphics/menu/stretch.gif"" width=""11"" height=""22"">"
	IF curtab=1 THEN
		response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
			"<a href=""/dgs/index.asp""><img src=""/shared/graphics/menu/dgs_on.gif"" width=""100"" height=""22"" border=""0""></a>"
	ELSE
		response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">" &_
			"<a href=""/dgs/index.asp""><img src=""/shared/graphics/menu/dgs_off.gif"" width=""100"" height=""22"" border=""0""></a>"
	END IF
	'-- tab2 --
	IF curtab=2 THEN
		response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
			"<a href=""/ok/list.asp?valid=0""><img src=""/shared/graphics/menu/ok_on.gif"" width=""100"" height=""22"" border=""0""></a>"
	ELSE
		IF curtab=1 THEN
			response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
		ELSE
			response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
		END IF
		response.write "<a href=""/ok/list.asp?valid=0""><img src=""/shared/graphics/menu/ok_off.gif"" width=""100"" height=""22"" border=""0""></a>"
	END IF
	'-- tab3 --
	IF curtab=3 THEN
		response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
			"<a href=""/dgb/list.asp?valid=0""><img src=""/shared/graphics/menu/dgb_on.gif"" width=""88"" height=""22"" border=""0""></a>"
	ELSE
		IF curtab=2 THEN
			response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
		ELSE
			response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
		END IF
		response.write "<a href=""/dgb/list.asp?valid=0""><img src=""/shared/graphics/menu/dgb_off.gif"" width=""88"" height=""22"" border=""0""></a>"
	END IF
	'-- tab 4 --
	
	IF curtab=4 THEN
		response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
			"<a href=""/kurser/list.asp?valid=0""><img src=""/shared/graphics/menu/kurser_on.gif"" width=""60"" height=""22"" border=""0""></a>"
	ELSE
		IF curtab=3 THEN
			response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
		ELSE
			response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
		END IF
		response.write "<a href=""/kurser/list.asp?valid=0""><img src=""/shared/graphics/menu/kurser_off.gif"" width=""60"" height=""22"" border=""0""></a>"
	END IF
	'-- tab 7 --
	
	IF curtab=7 THEN
		response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
			"<a href=""/art/list.asp?valid=0""><img src=""/shared/graphics/menu/artikler_on.gif"" width=""60"" height=""22"" border=""0""></a>"
	ELSE
		IF curtab=4 THEN
			response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
		ELSE
			response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
		END IF
		response.write "<a href=""/art/list.asp?valid=0""><img src=""/shared/graphics/menu/artikler_off.gif"" width=""60"" height=""22"" border=""0""></a>"
	END IF
	'-- tab5 --
	IF curtab=5 THEN
		response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
			"<a href=""/news/index.asp""><img src=""/shared/graphics/menu/news_on.gif"" width=""57"" height=""22"" border=""0""></a>"
	ELSE
		IF curtab=7 THEN
			response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
		ELSE
			response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
		END IF
		response.write "<a href=""/news/index.asp""><img src=""/shared/graphics/menu/news_off.gif"" width=""57"" height=""22"" border=""0""></a>"
	END IF
	'-- tab6 --
	IF curtab=6 THEN
		response.write "<img src=""/shared/graphics/menu/separator_left.gif"" width=""29"" height=""22"">" &_
			"<a href=""/admin/ei/tema/functions.asp""><img src=""/shared/graphics/menu/adm_on.gif"" width=""31"" height=""22"" border=""0""></a>" &_
			"<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
	ELSE
		IF curtab=5 THEN
			response.write "<img src=""/shared/graphics/menu/separator_right.gif"" width=""29"" height=""22"">"
		ELSE
			response.write "<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
		END IF
		response.write "<a href=""/admin/ei/tema/functions.asp""><img src=""/shared/graphics/menu/adm_off.gif"" width=""31"" height=""22"" border=""0""></a>" &_
			"<img src=""/shared/graphics/menu/separator.gif"" width=""29"" height=""22"">"
	END IF
	response.write "<img src=""/shared/graphics/menu/stretch.gif"" width=""11"" height=""22""></td>"
END IF
%>
</tr>
<%IF curtab>0 THEN%>
<tr><td class="menuBar">
<%
SET fs = CreateObject("Scripting.FileSystemObject")
Set ts=fs.OpenTextFile(Server.mapPath("inc_submenu.txt"))
execute(ts.ReadAll())
ts=""
fs=""
%>
</td></tr>
<%END IF%>
<tr><td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr></table><!-- #EndLibraryItem --><!-- #EndEditable --> 
<table width="750" border="0" cellspacing="0" cellpadding="0" name="Contentarea">
<tr> 
<td width="180" valign="top" name="leftbar"><!-- #BeginEditable "leftbar" -->
<table width="180" border="0" cellspacing="0" cellpadding="0">
<tr>
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" height="22" class="sidebarHeader">&nbsp;&nbsp;S&Aring;DAN G&Oslash;R DU</td>
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
<td valign="top">
<p>Tilmeld dig De gr&oslash;nne Sider ved at udfylde formularen. Brug de felter 
du finder n&oslash;dvendigt, men v&aelig;r opm&aelig;rksom p&aring;, at felter 
markeret med en * skal udfyldes.<br>
<br>
<span class="sidebarHeader">Navne</span><br>
Hvis relevant indtastes navnet p&aring; organisationen / firmaet / projeket.<br>
Vi skal altid have en kontakt-person.<br>
<span class="sidebarHeader"><br>
Kort og uddybende beskrivelse:</span><br>
Den korte beskrivelse vises i oversigten, mens kun den uddybende vises, n&aring;r 
nogen ser p&aring; dit kort i De Gr&oslash;nne Sider.</p>
<p><span class="sidebarHeader">Forsk&oslash;n teksten med HTML-formatering.</span> 
<br>
Brug Formater-knappen<br>
MAC brugere:<br>
<a href="#" onClick="MM_openBrWindow('/shared/HTMLhelp.htm','HTMLhj&aelig;lp','scrollbars=yes,width=800,height=600')">For 
MAC brugere.</a><br>
<br>
<span class="sidebarHeader">Kategori og emneord:</span><br>
- er vigtige for at du bliver vist rette sted. Dine valg her er kun vejledende 
for vores redakt&oslash;rer der foretager den endelige kategorisering.<br>
<br>
<span class="sidebarHeader">Stikord</span><br>
P&aring; &Oslash;ko-info kan man lave fritekst-s&oslash;gninger. Hvis du &oslash;nsker 
at blive vist n&aring;r der s&oslash;ges p&aring; ord der ikke findes i dine beskrivelser, 
kan de skrives her.<br>
<br>
<span class="sidebarHeader">Brugernavn og adgangskode</span>:<br>
- de oplysninger du (og andre i din organisation) skal logge ind med for at bruge 
&Oslash;ko-info Insider.<br>
Den email-adresse der indtastes her, er den vi sender til, n&aring;r vi har brug 
for at kontakte dig/Jer.<br>
<br>
</p>
</td>
</tr>
<tr>
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</table>
</td>
</tr>
</table>
<!-- #EndEditable --></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="569" valign="top" height="100%" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> <!-- #BeginEditable "maincontent" -->
<table width="100%" border="0" cellspacing="0" cellpadding="16" background="/dgs/graphics/sub_index_header_dgs_bkgrd.gif" name="subIndexHeader">
<form method="post" action="" onSubmit="donew(document.forms[0]);return document.validation" name="">
<tr> 
<td valign="top"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
<tr> 
<td class="contentHeader1">
<%IF LEN(request.cookies("eouserid"))=0 THEN 
	response.write "Tilmelding til De Gr&oslash;nne Sider"
ELSE
	response.write "Ny organisation"
END IF%>
</td>
</tr>
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr valign="top">
<td>
Organisation / Firma / Projekt<br>
<span class="formLabeltext">Navn:</span><br>
<input type="text" name="name" value="" size="32" class="formInputobjectLong">
<br>
<br>
</td>
<td>
Kontaktperson <br>
<span class="formLabeltext">Titel:</span><br>
<input type="text" name="title" value="" size="32" class="formInputobjectLong">
<span class="formLabeltext"><br>
Fornavn:</span><br>
<input type="text" name="firstname" value="" size="32" class="formInputobjectLong">
<span class="formLabeltext"><br>
Efternavn:</span><br>
<input type="text" name="lastname" value="" size="32" class="formInputobjectLong">
</td>
</tr>
<tr valign="top">
<td><br>
Adresse<br>
<span class="formLabeltext"> C/O (obs. C/O er automatisk foran):</span>&nbsp;&nbsp; 
<br>
<input type="text" name="adrco" value="" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext"> *Gade:</span>&nbsp;&nbsp; <br>
<input type="text" name="street1" value="" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext">Sted:&nbsp;&nbsp;</span> <br>
<input type="text" name="street2" value="" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext">*Postnr:</span>&nbsp;&nbsp; (ikke by)<br>
<input type="text" name="zip" value="" size="32" class="formInputobjectLong" onChange="javascript:checkPostnr(this.value)">
<input type="hidden" name="zip_dk">
</td>
<td><p><br>
Andre kontaktmuligheder <br>
<span class="formLabeltext"> Ved flere numre skriv da gerne lidt om nummeret i 
parentes: F.eks. xx xx xx xx (direkte/privat/mobil)<br>
Tlf 1:</span><br>
<input type="text" name="phone1" value="" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext">Tlf 2:</span><br>
<input type="text" name="phone2" value="" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext">Tlf 3:</span><br>
<input name="phone3" type="text" class="formInputobjectLong" id="phone3" value="" size="32">
<br>
<span class="formLabeltext">Fax:</span><br>
<input type="text" name="fax" value="" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext">*Email addresse 1:</span><br>
<input type="text" name="emailaddress1" value="" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext">Email addresse 2:</span><br>
<input type="text" name="emailaddress2" value="" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext">www:</span> <br>
(www.hjemmeside.dk - skriv ikke http://)<br>
<input type="text" name="www" value="" size="32" class="formInputobjectLong">
<br>
www2: <br>
<input type="text" name="www2" value="" size="32" class="formInputobjectLong">
</p></td>
</tr>
</table>
Beskrivelser
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr valign="top">
<td width="50%"> <span class="formLabeltext">Kort beskrivelse:</span><br>
<textarea name="shortdescription" cols="50" rows="5" class="formTextobjectLow"></textarea>
<div align="center"></div>
</td>
<td> 
<span class="formLabeltext">*Uddybende beskrivelse:</span><br>
<textarea name="description" cols="50" rows="5" class="formTextobjectBig"></textarea>
<div align="center"> 
<input type="button" class="function" onClick="showeditor('description','');" value="Formater" name="button">
</div>
</td>
</tr>
</table>
*Kategorier
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr valign="top">
<td width="50%">
<span class="formLabeltext">Alle kategorier:</span><br>
<select name="allCat" class="formTextobjectLow" size="5" ondblclick="addValue(this.form.allCat,this.form.selCat);">
<script src="/log/insider/ei/menufiles/dgscat_options.js"></script>
</select>
<br>
Dobbeltklik p&aring; en kategori for at v&aelig;lge den

</td>
<td>
<span class="formLabeltext">Valgte kategorier:</span><br>
<select name="selCat" class="formTextobjectLow" size="5" ondblclick="removeValue(this.form.selCat);" multiple>
</select>
<br>
Dobbeltklik p&aring; en kategori for at fjerne den

</td>
</tr>
</table>
<br>
*Emneord
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr valign="top">
<td width="50%">
<span class="formLabeltext">Alle emneord:</span><br>
<select name="allSbj" class="formTextobjectLow" size="5"  ondblclick="addValue(this.form.allSbj,this.form.selSbj);">
<script src="/log/insider/ei/menufiles/sbj_options.js"></script>
</select>
<br>
Dobbeltklik p&aring; et emneord for at v&aelig;lge det
</td>
<td>
<span class="formLabeltext">Valgte Emneord:</span><br>
<select name="selSbj" class="formTextobjectLow" size="5" ondblclick="removeValue(this.form.selSbj);" multiple>
</select>
<br>
Dobbeltklik p&aring; et emneord for at fjerne det
</td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr valign="top">
<td width="50%">
<span class="formLabeltext"><br>
Stikord:</span><br>
<textarea name="keywords" class="formTextobjectLow"></textarea>
</td>
<td>
<span class="formLabeltext"><br>
*Brugernavn:</span><br>
<input type="text" name="username" value="" size="32" class="formInputobjectLong">
<span class="formLabeltext"><br>
*Ønsket Adgangskode:</span><br>
<input type="text" name="password" value="" size="32" class="formInputobjectLong">
<span class="formLabeltext"><br>
*Email-adresse:</span><br>
<input type="text" name="email" value="" size="32" class="formInputobjectLong">
</td>
</tr>
</table>
<div align="center"><br>
<!--#include virtual="/shared/inc_getparams.asp" -->
<%=GetParams("&orgid=")%>
<input type="reset" value="Ryd" class="formbutton">
<input type="submit" name="Submit" value="Opret" class="formSubmitbutton">
</div>
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
<!-- #EndEditable --> </td>
</tr>
<tr> 
<td valign="bottom" align="left"> 
<!--#include virtual="/shared/pagetools.asp"-->
</td>
</tr>
</table>
</td>
</tr>
</table>
</td>
<td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
</tr>
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="3"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr align="center"> 
<td colspan="3" class="footer" height="20" valign="middle">
<!--#include virtual="/shared/footer.asp"-->
</td>
</tr>
</table>
</body>
<!-- #EndTemplate --></html>
<!--#include virtual="/shared/log.asp"-->


