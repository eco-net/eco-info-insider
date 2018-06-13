<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include file="lib.asp"-->
<!--#include virtual="/shared/common.asp" -->
<%if request("opdater")<>"" then
DoUpdate()
end if
%>
<%
Dim rsPageData__theID
rsPageData__theID = "0"
if (request("id") <> "") then rsPageData__theID = request("id")

%>
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT *  FROM bu_Artikel INNER JOIN bu_Artikel_site ON bu_Artikel.Artikel_ID = bu_Artikel_site.artikel_id  WHERE bu_Artikel.Artikel_ID = " + Replace(rsPageData__theID, "'", "''") + ""
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>
<%
set rsCat = Server.CreateObject("ADODB.Recordset")
rsCat.ActiveConnection = MM_ecoinfo_STRING
rsCat.Source = "SELECT *  FROM bu_ArtikelCategorys  ORDER BY Name"
rsCat.CursorType = 0
rsCat.CursorLocation = 2
rsCat.LockType = 3
rsCat.Open()
rsCat_numRows = 0
%>
<%
'set rsTema = Server.CreateObject("ADODB.Recordset")
'rsTema.ActiveConnection = MM_ecoinfo_STRING
'rsTema.Source = "SELECT id, name  FROM eitema_maindata"
'rsTema.CursorType = 0
'rsTema.CursorLocation = 2
'rsTema.LockType = 3
'rsTema.Open()
'rsTema_numRows = 0
%>

<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Ecoinfo</title>
<script language="javascript">
<!--
function deleteNews()
{
if(confirm("Vil du slette nyheden?"))
{
	document.nyhedsform.sletNyhed.value="slet";
	document.nyhedsform.submit();
}
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') {
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (val<min || max<val) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<!-- #EndEditable --> 
<script src="/shared/common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
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
curtab=5
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
<td width="98%" height="22" class="sidebarHeader">&nbsp;&nbsp;S&Aring;DAN G&Oslash;R 
DU</td>
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
<p>&AElig;ndr de data du &oslash;nsker at &aelig;ndre, men v&aelig;r opm&aelig;rksom 
p&aring;, at felter markeret med en * skal v&aelig;re udfyldt.</p>
<p><span class="sidebarHeader">Fejlmuligheder<br>
</span>Lange tekster kopieres ind via Notesblok, og ikke fra Word, som har egne 
koder.<br>
Der kan ikke <br>
bruges: &quot; eller: ' .<br>
Bliv f&aelig;rdig inden 20 min., ellers lukkes Sessionen.<br>
<br>
<span class="sidebarHeader">Titel</span><br>
Ved visning af nyhed bruges titel som link. B&oslash;r v&aelig;re kort.</p>
<p><span class="sidebarHeader">Godkend</span><br>
Kun godkendte vises p&aring; nettet. Ikke godkendte placeres til godkendelse i 
Insider under Nyheder </p>
<p><span class="sidebarHeader">Kort og uddybende beskrivelse</span><br>
Den korte beskrivelse vises i oversigten, mens kun uddybende vises, n&aring;r 
&quot;hele&quot; nyheden vises. Gentagelser er alts&aring; tilladt.</p>
<p><span class="sidebarHeader">Forsk&oslash;n teksten med HTML-formatering.</span> 
<br>
<a href="#" onClick="MM_openBrWindow('/shared/HTMLhelp.htm','HTMLhj&aelig;lp','scrollbars=yes,width=800,height=600')">Se 
her hvordan.</a></p>
<p><span class="sidebarHeader">Link</span><br>
er en http adresse, som enten vises i et nyt vindue, eller i det samme vindue. 
Skriv <span class="listheader">http://www.adresse.dk</span><br>
<br>
<span class="sidebarHeader">Linktekst <br>
</span>er det link man klikker p&aring;.<br>
Undg&aring; lange tekster med linieskift: &lt;br&gt; </p>
<p><span class="sidebarHeader">Vises p&aring; site:<br>
</span>Afm&aelig;rkes for en eller begge sites. Bliver altid vist p&aring; &Oslash;ko-info</p>
<p><span class="sidebarHeader">Billede:</span><br>
Billeder i Billedbasen kan vises i h&oslash;jre kolonne af detail siden. Skriv 
blot billednummer.</p>
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
<tr> 
<td valign="top"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
<tr> 
<td class="contentHeader1"> Rediger Nyhed Mac</td>
</tr>
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top"> 
<%'IF LEN(request.cookies("eiuserid"))>0 THEN
	'response.write "Tilknyttet: "
	'IF LEN(rsOrg.Fields.Item("name").value)>0 THEN
	'	response.write rsOrg.Fields.Item("name").value & "<br><br>"
	'ELSE
	'	response.write rsOrg.Fields.Item("firstname").value & " " & rsOrg.Fields.Item("lastname").value & "<br><br>"
	'END IF
'END IF%>
<br>
<form name="nyhedsform" method="post" action="edit_mac.asp">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr valign="top"> 
<td width="50%"> 
<p><span class="formLabeltext">*Titel:</span><br>
<input value="<%=toString(rsPageData.Fields.Item("Header").Value)%>" type="text" name="title" size="32" class="formInputobjectLong">
<br>
<span class="listheader">*Tilknyt kategori:</span><br>
<select name="cat" class="formInputobjectLong">
<!--cat_id 23 = Bæredygtig udvikling - generelt -->
<option value="23" selected>Vælg Kategori</option>
<%
While (NOT rsCat.EOF)
%>
<% if (rsCat.Fields.Item("ArtikelCategory_ID").Value)=(rsPageData.Fields.Item("ArtikelCategory_ID").Value) then %>
<option value="<%=(rsCat.Fields.Item("ArtikelCategory_ID").Value)%>" selected><%=(rsCat.Fields.Item("Name").Value)%> 
<%else %>
<option value="<%=(rsCat.Fields.Item("ArtikelCategory_ID").Value)%>" ><%=(rsCat.Fields.Item("Name").Value)%> 
<%end if %>
</option>
<%
  rsCat.MoveNext()
Wend
If (rsCat.CursorType > 0) Then
  rsCat.MoveFirst
Else
  rsCat.Requery
End If
%>
</select>
</p>
<p>&nbsp; </p>
<p> <br>
</p>
<p><span class="formLabeltext"> </span></p></td>
<td width="50%"> 
<p><span class="formLabeltext">*Kort beskrivelse:</span><br>
<textarea name="shortdescription" cols="50" rows="5" class="formTextobjectLow"><%=toString(rsPageData.Fields.Item("ShortDescription").Value)%></textarea>
</p></td>
</tr>
<tr valign="top"> 
<td width="50%"><span class="formLabeltext">*Uddybende beskrivelse:</span> <span class="formLabeltext"><br>
<textarea name="EditorValue" class="formTextobjectBig"><%=toString(rsPageData.Fields.Item("Content").Value)%></textarea>
</span></td>
<td width="50%"> 
<p>&nbsp;</p>
<p align="center"><span class="sidebarHeader">Forsk&oslash;n teksten med HTML-formatering.</span> 
<br>
<a href="/shared/HTMLhelp.htm" target="_blank">Se her hvordan.</a></p>
<p><span class="sidebarHeader"></span></p></td>
</tr>
<tr valign="top">
  <td>&nbsp;</td>
  <td>&nbsp;</td>
</tr>
<tr valign="top">
  <td><table width="100%" border="0" cellpadding="0">
    <tr>
      <td><span class="listheader">Link1:</span></td>
      <td><input name="link" type="text" class="formInputobject" id="link" value="<%=(rsPageData.Fields.Item("FileName").Value)%>" size="32"></td>
    </tr>
    <tr>
      <td><span class="listheader">Linktekst:</span></td>
      <td><input name="linktext" type="text" class="formInputobject" id="linktext" value="<%=(rsPageData.Fields.Item("linktext").Value)%>"></td>
    </tr>
    <tr>
      <td><span class="listheader">Link2:</span></td>
      <td><input name="link2" type="text" class="formInputobject" id="link2" value="<%=(rsPageData.Fields.Item("FileName2").Value)%>" size="32"></td>
    </tr>
    <tr>
      <td><span class="listheader">Linktekst:</span></td>
      <td><input name="linktext2" type="text" class="formInputobject" id="linktext2" value="<%=(rsPageData.Fields.Item("linktext2").Value)%>"></td>
    </tr>
    <tr>
      <td><span class="listheader">Link3:</span></td>
      <td><input name="link3" type="text" class="formInputobject" id="link3" value="<%=(rsPageData.Fields.Item("FileName3").Value)%>" size="32"></td>
    </tr>
    <tr>
      <td><span class="listheader">Linktekst:</span></td>
      <td><input name="linktext3" type="text" class="formInputobject" id="linktext3" value="<%=(rsPageData.Fields.Item("linktext3").Value)%>"></td>
    </tr>
    <tr>
      <td><span class="listheader">Link4:</span></td>
      <td><input name="link4" type="text" class="formInputobject" id="link4" value="<%=(rsPageData.Fields.Item("link4").Value)%>" size="32"></td>
    </tr>
    <tr>
      <td><span class="listheader">Linktekst:</span></td>
      <td><input name="linktext4" type="text" class="formInputobject" id="linktext4" value="<%=(rsPageData.Fields.Item("linktext4").Value)%>"></td>
    </tr>
    <tr>
      <td><div align="right"><span class="listheader">
          <input type="radio" name="vindue" <%if ((rsPageData.Fields.Item("newwin").Value)=0) then response.write("checked")%> value="0">
      </span></div></td>
      <td><span class="listheader">&Aring;bnes i nyt vindue </span></td>
    </tr>
    <tr>
      <td><div align="right"><span class="listheader">
          <input type="radio" name="vindue" <%if ((rsPageData.Fields.Item("newwin").Value)=1) then response.write("checked")%> value="1">
      </span></div></td>
      <td><span class="listheader">&Aring;bnes i samme vindue</span></td>
    </tr>
  </table></td>
  <td><p><span class="listheader">Vises p&aring; site:</span><br>
        <input name="bu" type="checkbox" id="bu" value="checkbox" <%if ((rsPageData.Fields.Item("busite").Value)=1) then response.write("checked")%>>
    baeredygtigudvikling.nu &nbsp;&nbsp;&nbsp;&nbsp; <br>
  <input name="econet" type="checkbox" id="econet" value="ok" <%if ((rsPageData.Fields.Item("econetsite").Value)=1) then response.write("checked")%>>
    eco-net.dk &nbsp;&nbsp;&nbsp;&nbsp; <br>
    Alle nyheder vises p&aring; eco-info.dk<br>
    p&aring;n&aelig;r eco-net, som kun vises p&aring; eco-net.dk</p>
    <p><span class="formLabeltext">
      <label>
      <input name="miljoinfo" type="checkbox" id="miljoinfo" value="checkbox" <%if rsPageData("miljoinfo")>0 then response.write "checked"%>>
      </label>
      Milj&oslash; Info vises </span></p>
    <p><span class="formLabeltext">
      <input name="validated" type="checkbox" id="validated" value="ok" <%if (rsPageData.Fields.Item("Approved").Value)=1 then response.write("checked")%>>
      Godkendt</span></p></td>
</tr>
<tr valign="top">
  <td>&nbsp;</td>
  <td>&nbsp;</td>
</tr>
<tr valign="top">
  <td><table width="100%" border="0" cellpadding="0">
    <tr>
      <td><strong><span class="listheader">Inds&aelig;t billede i midten, under tekst</span></strong></td>
    </tr>
    <tr>
      <td><a href="/shared/picasa.asp" target="_blank"><strong>Billede2 fra Picasa</strong></a><a href="/shared/picasa.asp"></a>:<br>
          <input name="imagefilename2" type="text" class="formInputobjectLong" id="imagefilename2" value="<%=rsPageData("imagefilename2")%>" size="40" /></td>
    </tr>
    <tr>
      <td>Billedtekst: <br>
          <input name="img_text2" type="text" class="formInputobject" id="img_text2" value="<%=rsPageData("img_text2")%>"></td>
    </tr>
  </table></td>
  <td><table width="100%" border="0" cellpadding="0">
    <tr>
      <td><span class="listheader">Inds&aelig;t billede i h&oslash;jre kolonne </span></td>
    </tr>
    <tr>
      <td><a href="/shared/picasa.asp" target="_blank"><strong>Billede1 fra Picasa</strong></a><a href="/shared/picasa.asp"></a>:<br>
          <input name="imagefilename1" type="text" class="formInputobjectLong" id="imagefilename1" value="<%=rsPageData("imagefilename1")%>" size="40" /></td>
    </tr>
    <tr>
      <td>Billedtekst:<br>
          <input name="img_text" type="text" class="formInputobject" id="img_text" value="<%=rsPageData("img_text")%>"></td>
    </tr>
  </table></td>
</tr>
<tr valign="top">
  <td>&nbsp;</td>
  <td>&nbsp;</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top">&nbsp;<!--TEMA<div align="center">Hvis nyheden tilh&oslash;rer et tema:<br>
  <select name="tema" class="formInputobject">
  <option selected value="0">vælg tema 
  <option> 
  <%
'While (NOT rsTema.EOF)
%>
  <option value="<%'=(rsTema.Fields.Item("id").Value)%>" <%'if rsPageData("Tema_ID")=rsTema("id") then response.write("selected")%>><%'=(rsTema.Fields.Item("name").Value)%> </option>
  <%
  'rsTema.MoveNext()
'Wend
'If (rsTema.CursorType > 0) Then
 ' rsTema.MoveFirst
'Else
  'rsTema.Requery
'End If
%>
  </select>
</div>TEMA-->
</td>
</tr>

<tr> 
<td> 
    <p align="center" class="formLabeltext">
      <label></label>
      <input type="submit" name="Submit" value="Gem" class="formSubmitbutton" onClick="MM_validateForm('title','','R','shortdescription','','R')">
      <input type="button" name="SubmitSlet" value="Slet" class="formSubmitbutton" onClick="deleteNews()">
      <input name="opdater" type="hidden" id="opdater" value="1">
          <input name="sletNyhed" type="hidden" id="sletNyhed">
          <input name="id" type="hidden" id="id" value="<%=request("id")%>">
          <input name="createDate" type="hidden" id="createDate" value="<%=CStr(Date)%>">
          <input name="tema" type="hidden" id="tema" value="0">
          </p>    </td>
</tr>
</table>
</form>
</td>
</tr>
</table>
</td>
</tr>
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
<%
rsPageData.Close()
%>
<%
rsCat.Close()
%>
<!--#include virtual="/shared/log.asp"-->

