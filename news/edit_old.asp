<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include file="lib.asp"-->
<!--#include virtual="/shared/common.asp" -->
<script language="javascript">
if (navigator.platform=="Mac68k" || navigator.platform=="MacPPC"){
document.write("platform: " + navigator.platform);
location.replace("edit_mac.asp?id=<%=request("id")%>")
}
</script>
<%if request("opdater")<>"" then
DoUpdate()
Response.redirect("index.asp")
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
set rsTema = Server.CreateObject("ADODB.Recordset")
rsTema.ActiveConnection = MM_ecoinfo_STRING
rsTema.Source = "SELECT id, name  FROM eitema_maindata"
rsTema.CursorType = 0
rsTema.CursorLocation = 2
rsTema.LockType = 3
rsTema.Open()
rsTema_numRows = 0
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
<p><span class="sidebarHeader">Link</span><br>
er en http adresse, som enten vises i et nyt vindue, eller i det samme vindue. 
Skriv <span class="listheader">http://www.adresse.dk</span><br>
<br>
<span class="sidebarHeader">Linktekst <br>
</span>er det link man klikker p&aring;.<br>
Undg&aring; lange tekster med linieskift: &lt;br&gt;</p>
<p><span class="sidebarHeader">Vises p&aring; site:<br>
</span>Afm&aelig;rkes for en eller begge sites. Bliver altid vist p&aring; &Oslash;ko-info</p>
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
<td class="contentHeader1"> Rediger Nyhed</td>
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
<form name="nyhedsform" method="post" action="edit.asp">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr valign="top"> 
<td> 
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
<p><span class="formLabeltext"> </span></p>
</td>
<td> 
<p><span class="formLabeltext">*Kort beskrivelse:</span><br>
<textarea name="shortdescription" cols="50" rows="5" class="formTextobjectLow"><%=toString(rsPageData.Fields.Item("ShortDescription").Value)%></textarea>
</p>
</td>
</tr>
<tr valign="top"> 
<td colspan="2"><span class="formLabeltext">*Uddybende beskrivelse:</span> 
<style type="text/css">
<!--
.clsCursor {  cursor: hand}
-->
</style>
<textarea name="EditorValue" style="display: none;"><%=toString(rsPageData.Fields.Item("Content").Value)%> </textarea>
<script language="JavaScript">

  var errorString = "Sorry but this web page needs\nWindows95 and Internet Explorer 5 or above to view."
  var Ok = "false";
  var name =  navigator.appName;
  var version =  parseFloat(navigator.appVersion);
  var platform = navigator.platform;

	if (platform == "Win32" && name == "Microsoft Internet Explorer" && version >= 4){
		Ok = "true";
	} else {
		Ok= "false";
	}

	if (Ok == "false") {
		alert(errorString);
	}

function ColorPalette_OnClick(colorString){
	
	cpick.bgColor=colorString;
	document.all.colourp.value=colorString;
	doFormat('ForeColor',colorString);
}

function initToolBar(ed) {
    
	var eb = document.all.editbar;
	if (ed!=null) {
		eb._editor = window.frames['myEditor'];
	}
}

function doFormat(what) {

	var eb = document.all.editbar;
		
	if(what == "FontName"){
		if(arguments[1] != 1){
			eb._editor.execCommand(what, arguments[1]);
			document.all.font.selectedIndex = 0;
		} 
	} else if(what == "FontSize"){
    if(arguments[1] != 1){
      eb._editor.execCommand(what, arguments[1]);
      document.all.size.selectedIndex = 0;
    } 
	} else {
	   eb._editor.execCommand(what, arguments[1]);
	}
}

function swapMode() {

	var eb = document.all.editbar._editor;
  eb.swapModes();
}

function create() {

    var eb = document.all.editbar;
    eb._editor.newDocument();
}

function newFile(){

	create();
}

function makeUrl(){

	sUrl = document.all.what.value + document.all.url.value;
	doFormat('CreateLink',sUrl);
}

function makeBlankLink(theHtml)
{
	a=/<A href/gi;
	t="<A target='_blank' href";
	return theHtml.replace(a, t);
}
function copyValue() {
	a=/<A href/gi;
	t="<A target=_blank href";
	var theHtml = "" + document.frames("myEditor").document.frames("textEdit").document.body.innerHTML + "";
	theString=theHtml.replace(a, t);
	document.all.EditorValue.value = theString;
}
function SwapView_OnClick(){

  if(document.all.btnSwapView.value == "View Html"){
		var sMes = "View Wysiwyg";
    var sStatusBarMes = "Current View Html";
	} else {
		var sMes = "View Html"
    var sStatusBarMes = "Current View Wysiwyg";
  }
	
	document.all.btnSwapView.value = sMes;
  window.status  = sStatusBarMes;
	swapMode();
}

function Help_OnClick(){
  window.open("editor_images/help_document.htm","wHelp", "toolbar=0, scrollbars=yes, width=640, height=480");
}

function OnFormSubmit(){
     copyValue();
}
</script>
<table border="1" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="60%" height="40%" bordercolor="#CCCCCC">
<tr valign="top"> 
<td> 
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
<tr valign="top"> 
<td valign="top"> 
<div id=editbar > 
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="left">
<tr> 
<td> 
<table border="0" cellpadding="0" cellspacing="0">
<tr> 
<td> 
<table border="0">
<tr valign="baseline"> 
<td nowrap> <img class='clsCursor' src="editor_images/Cut.gif" width="16" height="16" border="0" alt="Cut " onClick="doFormat('Cut')">&nbsp 
<img class='clsCursor' src="editor_images/Copy.gif" width="16" height="16" border="0" alt="Copy" onClick="doFormat('Copy')">&nbsp 
<img class='clsCursor' src="editor_images/Paste.gif" border="0" alt="Paste" onClick="doFormat('Paste')" width="16" height="16">&nbsp 
</td>
</tr>
</table>
</td>
<td> 
<table border="0">
<tr valign="baseline"> 
<td nowrap> <img class='clsCursor' src="editor_images/para_bul.gif" width="16" height="16" border="0" alt="Bullet List" onClick="doFormat('InsertUnorderedList');" >&nbsp 
<img class='clsCursor' src="editor_images/para_num.gif" width="16" height="16" border="0" alt="Numbered List" onClick="doFormat('InsertOrderedList');" >&nbsp 
<img class='clsCursor' src="editor_images/indent.gif" width="20" height="16" alt="Indent" onClick="doFormat('Indent')">&nbsp 
<img class='clsCursor' src="editor_images/outdent.gif" width="20" height="16" alt="Outdent" onClick="doFormat('Outdent')">&nbsp 
<img class='clsCursor' src="editor_images/hr.gif" width="16" height="18" alt="HR" onClick="doFormat('InsertHorizontalRule')">&nbsp 
</td>
</tr>
</table>
</td>
<td> 
<table border="0">
<tr valign="baseline"> 
<td nowrap> 
<select name="what" style="font: 8pt verdana;">
<option value="http://" selected>http://</option>
<option value="mailto:">mailto:</option>
<option value="ftp://">ftp://</option>
<option value="https://">https://</option>
</select>
</td>
<td> 
<input type="text" name="url" size="30" style="font: 8pt verdana;">
</td>
<td> 
<input type="button" name="button2" value="Add" onClick="makeUrl();" style="font: 8pt verdana;">
</td>
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
<tr> 
<td height="41"> 
<table border="0" width="376">
<tr> 
<td nowrap valign="baseline"> 
<div align="left"> 
<select name="font" onChange=" doFormat('FontName',document.all.font.value);" style="font: 8pt verdana;">
<option value="1" selected >Select Font...</option>
<option value="arial">Arial, Helvetica, sans-serif</option>
<option value="times" >Times New Roman, Times, serif</option>
<option value="courier">Courier New, Courier, mono</option>
<option value="georgia">Georgia, Times New Roman</option>
<option value="verdana">Verdana, Arial, Helvetica</option>
</select>
<select name="size" onChange="doFormat('FontSize',document.all.size.value);" style="font: 8pt verdana;">
<option value="None" selected>Size</option>
<option value="1">1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
<option value="6">6</option>
<option value="7">7</option>
<option value="+1">+1</option>
<option value="+2">+2</option>
<option value="+3">+3</option>
<option value="+4">+4</option>
<option value="+5">+5</option>
<option value="+6">+6</option>
<option value="+7">+7</option>
</select>
<img class='clsCursor' src="editor_images/Bold.gif" width="16" height="16" border="0" align="absmiddle" alt="Bold text" onClick="doFormat('Bold')">&nbsp 
<img class='clsCursor' src="editor_images/Italics.gif" width="16" height="16" border="0" align="absmiddle" alt="Italic text" onClick="doFormat('Italic')">&nbsp 
<img class='clsCursor' src="editor_images/underline.gif" width="16" height="16" border="0" align="absmiddle" alt="Underline text" onClick="doFormat('Underline')" >&nbsp 
<img class='clsCursor' src="editor_images/left.gif" width="16" height="16" border="0" alt="Align Left" align="absmiddle"  onClick="doFormat('JustifyLeft')"> 
<img class='clsCursor' src="editor_images/centre.gif" width="16" height="16" border="0" alt="Align Center" align="absmiddle" onClick="doFormat('JustifyCenter')">&nbsp 
<img class='clsCursor' src="editor_images/right.gif" width="16" height="16" border="0" alt="Align Right" align="absmiddle"  onClick="doFormat('JustifyRight')">&nbsp 
</div>
</td>
<td><img class='clsCursor' src="editor_images/help.gif" width="20" height="20" align="middle" alt="Help" onClick="Help_OnClick();"></td>
</tr>
</table>
</td>
</tr>
</table>
</div>
</td>
</tr>
<tr valign="top" align="left"> 
<td valign="top"> 
<table width="100%" border="0" height="100%">
<tr valign="top"> 
<td width="90%" height="90%"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
<tr valign="top"> 
<td bgcolor="#FFFFFF"><iframe id=myEditor src="pd_edit.htm" onFocus="initToolBar(this)" width="100%" height="100%"></iframe></td>
</tr>
</table>
</td>
<td width="9%" align="center"> 
<table  bgcolor="#000000" width="74" id="cpick" border="1" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td>&nbsp;</td>
</tr>
</table>
<input type="text" name="colourp" size="8" value="#000000" style="width:74px; font: 8pt verdana" readonly>
<table border=1 bgcolor="#CCCCCC" cellpadding="0" cellspacing="0" width="74">
<tr> 
<td bgcolor="#ffffff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ffffff')"></td>
<td bgcolor="#ffffcc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ffffcc')"></td>
<td bgcolor="#ffff99" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ffff99')"></td>
<td bgcolor="#ffff66" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ffff66')"></td>
<td bgcolor="#ffff33" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ffff33')"></td>
<td bgcolor="#ffff00" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ffff00')"></td>
</tr>
<tr> 
<td bgcolor="#ccffff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ccffff')"></td>
<td bgcolor="#ccffcc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ccffcc')"></td>
<td bgcolor="#ccff99" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ccff99')"></td>
<td bgcolor="#ccff66" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ccff66')"></td>
<td bgcolor="#ccff33" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ccff33')"></td>
<td bgcolor="#ccff00" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ccff00')"></td>
</tr>
<tr> 
<td bgcolor="#99ffff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#99ffff')"></td>
<td bgcolor="#99ffcc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#99ffcc')"></td>
<td bgcolor="#99ff99" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#99ff99')"></td>
<td bgcolor="#99ff66" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#99ff66')"></td>
<td bgcolor="#99ff33" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#99ff33')"></td>
<td bgcolor="#99ff00" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#99ff00')"></td>
</tr>
<tr> 
<td bgcolor="#00ffff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#00ffff')"></td>
<td bgcolor="#00ffcc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#00ffcc')"></td>
<td bgcolor="#00ff99" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#00ff99')"></td>
<td bgcolor="#00ff66" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#00ff66')"></td>
<td bgcolor="#00ff33" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#00ff33')"></td>
<td bgcolor="#00ff00" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#00ff00')"></td>
</tr>
<tr> 
<td bgcolor="#ffccff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ffccff')"></td>
<td bgcolor="#ffcccc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ffcccc')"></td>
<td bgcolor="#ffcc99" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ffcc99')"></td>
<td bgcolor="#ffcc66" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ffcc66')"></td>
<td bgcolor="#ffcc33" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ffcc33')"></td>
<td bgcolor="#ffcc00" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ffcc00')"></td>
</tr>
<tr> 
<td bgcolor="#ccccff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ccccff')"></td>
<td bgcolor="#cccccc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cccccc')"></td>
<td bgcolor="#cccc99" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cccc99')"></td>
<td bgcolor="#cccc66" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cccc66')"></td>
<td bgcolor="#cccc33" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cccc33')"></td>
<td bgcolor="#cccc00" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cccc00')"></td>
</tr>
<tr> 
<td bgcolor="#00ccff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#00ccff')"></td>
<td bgcolor="#00cccc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#00cccc')"></td>
<td bgcolor="#00cc99" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#00cc99')"></td>
<td bgcolor="#00cc66" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#00cc66')"></td>
<td bgcolor="#00cc33" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#00cc33')"></td>
<td bgcolor="#00cc00" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#00cc00')"></td>
</tr>
<tr> 
<td bgcolor="#ff99ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff99ff')"></td>
<td bgcolor="#ff99cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff99cc')"></td>
<td bgcolor="#ff9999" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff9999')"></td>
<td bgcolor="#ff9966" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff9966')"></td>
<td bgcolor="#ff9933" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff9933')"></td>
<td bgcolor="#ff9900" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff9900')"></td>
</tr>
<tr> 
<td bgcolor="#cc99ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc99ff')"></td>
<td bgcolor="#cc99cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc99cc')"></td>
<td bgcolor="#cc9999" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc9999')"></td>
<td bgcolor="#cc9966" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc9966')"></td>
<td bgcolor="#cc9933" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc9933')"></td>
<td bgcolor="#cc9900" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc9900')"></td>
</tr>
<tr> 
<td bgcolor="#9999ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#9999ff')"></td>
<td bgcolor="#9999cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#9999cc')"></td>
<td bgcolor="#999999" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#999999')"></td>
<td bgcolor="#999966" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#999966')"></td>
<td bgcolor="#999933" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#999933')"></td>
<td bgcolor="#999900" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#999900')"></td>
</tr>
<tr> 
<td bgcolor="#6699ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#6699ff')"></td>
<td bgcolor="#6699cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#6699cc')"></td>
<td bgcolor="#669999" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#669999')"></td>
<td bgcolor="#669966" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#669966')"></td>
<td bgcolor="#669933" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#669933')"></td>
<td bgcolor="#669900" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#669900')"></td>
</tr>
<tr> 
<td bgcolor="#3399ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#3399ff')"></td>
<td bgcolor="#3399cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#3399cc')"></td>
<td bgcolor="#339999" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#339999')"></td>
<td bgcolor="#339966" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#339966')"></td>
<td bgcolor="#339933" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#339933')"></td>
<td bgcolor="#339900" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#339900')"></td>
</tr>
<tr> 
<td bgcolor="#0099ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#0099ff')"></td>
<td bgcolor="#0099cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#0099cc')"></td>
<td bgcolor="#009999" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#009999')"></td>
<td bgcolor="#009966" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#009966')"></td>
<td bgcolor="#009933" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#009933')"></td>
<td bgcolor="#009900" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#009900')"></td>
</tr>
<tr> 
<td bgcolor="#ff66ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff66ff')"></td>
<td bgcolor="#ff66cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff66cc')"></td>
<td bgcolor="#ff6699" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff6699')"></td>
<td bgcolor="#ff6666" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff6666')"></td>
<td bgcolor="#ff6633" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff6633')"></td>
<td bgcolor="#ff6600" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff6600')"></td>
</tr>
<tr> 
<td bgcolor="#cc66ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc66ff')"></td>
<td bgcolor="#cc66cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc66cc')"></td>
<td bgcolor="#cc6699" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc6699')"></td>
<td bgcolor="#cc6666" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc6666')"></td>
<td bgcolor="#cc6633" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc6633')"></td>
<td bgcolor="#cc6600" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc6600')"></td>
</tr>
<tr> 
<td bgcolor="#9966ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#9966ff')"></td>
<td bgcolor="#9966cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#9966cc')"></td>
<td bgcolor="#996699" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#996699')"></td>
<td bgcolor="#996666" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#996666')"></td>
<td bgcolor="#996633" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#996633')"></td>
<td bgcolor="#996600" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#996600')"></td>
</tr>
<tr> 
<td bgcolor="#6666ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#6666ff')"></td>
<td bgcolor="#6666cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#6666cc')"></td>
<td bgcolor="#666699" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#666699')"></td>
<td bgcolor="#666666" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#666666')"></td>
<td bgcolor="#666633" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#666633')"></td>
<td bgcolor="#666600" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#666600')"></td>
</tr>
<tr> 
<td bgcolor="#3366ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#3366ff')"></td>
<td bgcolor="#3366cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#3366cc')"></td>
<td bgcolor="#336699" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#336699')"></td>
<td bgcolor="#336666" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#336666')"></td>
<td bgcolor="#336633" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#336633')"></td>
<td bgcolor="#336600" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#336600')"></td>
</tr>
<tr> 
<td bgcolor="#0066ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#0066ff')"></td>
<td bgcolor="#0066cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#0066cc')"></td>
<td bgcolor="#006699" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#006699')"></td>
<td bgcolor="#006666" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#006666')"></td>
<td bgcolor="#006633" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#006633')"></td>
<td bgcolor="#006600" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#006600')"></td>
</tr>
<tr> 
<td bgcolor="#ff33ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff33ff')"></td>
<td bgcolor="#ff33cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff33cc')"></td>
<td bgcolor="#ff3399" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff3399')"></td>
<td bgcolor="#ff3366" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff3366')"></td>
<td bgcolor="#ff3333" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff3333')"></td>
<td bgcolor="#ff3300" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff3300')"></td>
</tr>
<tr> 
<td bgcolor="#cc33ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc33ff')"></td>
<td bgcolor="#cc33cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc33cc')"></td>
<td bgcolor="#cc3399" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc3399')"></td>
<td bgcolor="#cc3366" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc3366')"></td>
<td bgcolor="#cc3333" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc3333')"></td>
<td bgcolor="#cc3300" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc3300')"></td>
</tr>
<tr> 
<td bgcolor="#9933ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#9933ff')"></td>
<td bgcolor="#9933cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#9933cc')"></td>
<td bgcolor="#993399" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#993399')"></td>
<td bgcolor="#993366" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#993366')"></td>
<td bgcolor="#993333" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#993333')"></td>
<td bgcolor="#993300" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#993300')"></td>
</tr>
<tr> 
<td bgcolor="#6633ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#6633ff')"></td>
<td bgcolor="#6633cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#6633cc')"></td>
<td bgcolor="#663399" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#663399')"></td>
<td bgcolor="#663366" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#663366')"></td>
<td bgcolor="#663333" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#663333')"></td>
<td bgcolor="#663300" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#663300')"></td>
</tr>
<tr> 
<td bgcolor="#3333ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#3333ff')"></td>
<td bgcolor="#3333cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#3333cc')"></td>
<td bgcolor="#333399" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#333399')"></td>
<td bgcolor="#333366" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#333366')"></td>
<td bgcolor="#333333" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#333333')"></td>
<td bgcolor="#333300" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#333300')"></td>
</tr>
<tr> 
<td bgcolor="#0033ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#0033ff')"></td>
<td bgcolor="#0033cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#0033cc')"></td>
<td bgcolor="#003399" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#003399')"></td>
<td bgcolor="#003366" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#003366')"></td>
<td bgcolor="#003333" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#003333')"></td>
<td bgcolor="#003300" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#003300')"></td>
</tr>
<tr> 
<td bgcolor="#ff00ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff00ff')"></td>
<td bgcolor="#ff00cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff00cc')"></td>
<td bgcolor="#ff0099" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff0099')"></td>
<td bgcolor="#ff0066" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff0066')"></td>
<td bgcolor="#ff0033" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff0033')"></td>
<td bgcolor="#ff0000" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#ff0000')"></td>
</tr>
<tr> 
<td bgcolor="#cc00ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc00ff')"></td>
<td bgcolor="#cc00cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc00cc')"></td>
<td bgcolor="#cc0099" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc0099')"></td>
<td bgcolor="#cc0066" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc0066')"></td>
<td bgcolor="#cc0033" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc0033')"></td>
<td bgcolor="#cc0000" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#cc0000')"></td>
</tr>
<tr> 
<td bgcolor="#9900ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#9900ff')"></td>
<td bgcolor="#9900cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#9900cc')"></td>
<td bgcolor="#990099" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#990099')"></td>
<td bgcolor="#990066" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#990066')"></td>
<td bgcolor="#990033" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#990033')"></td>
<td bgcolor="#990000" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#990000')"></td>
</tr>
<tr> 
<td bgcolor="#6600ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#6600ff')"></td>
<td bgcolor="#6600cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#6600cc')"></td>
<td bgcolor="#660099" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#660099')"></td>
<td bgcolor="#660066" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#660066')"></td>
<td bgcolor="#660033" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#660033')"></td>
<td bgcolor="#660000" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#660000')"></td>
</tr>
<tr> 
<td bgcolor="#3300ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#3300ff')"></td>
<td bgcolor="#3300cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#3300cc')"></td>
<td bgcolor="#330099" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#330099')"></td>
<td bgcolor="#330066" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#330066')"></td>
<td bgcolor="#330033" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#330033')"></td>
<td bgcolor="#330000" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#330000')"></td>
</tr>
<tr> 
<td bgcolor="#0000ff" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#0000ff')"></td>
<td bgcolor="#0000cc" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#0000cc')"></td>
<td bgcolor="#000099" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#000099')"></td>
<td bgcolor="#000066" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#000066')"></td>
<td bgcolor="#000033" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#000033')"></td>
<td bgcolor="#000000" width="12"><img class="clsCursor" src="blank.gif" height=8 width=10 border=0 onClick="ColorPalette_OnClick('#000000')"></td>
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
<script>
  initToolBar("foo");
  window.status  = "Current View: Wysiwyg";
</script>
</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td width="30%" bgcolor="#CCCCCC"> <span class="listheader"> &nbsp;Linktekst:</span> 
<input type="text" name="linktext" value="<%=(rsPageData.Fields.Item("linktext").Value)%>" class="formInputobject">
</td>
<td width="30%" valign="top" bgcolor="#CCCCCC"> <span class="listheader">Link:</span> 
<input value="<%=(rsPageData.Fields.Item("FileName").Value)%>" type="text" name="link" size="32" class="formInputobject">
</td>
<td width="40%" rowspan="3"> <span class="listheader">Vises p&aring; site:</span><br>
<input type="checkbox" name="bu" value="checkbox" <%if ((rsPageData.Fields.Item("busite").Value)=1) then response.write("checked")%>>
baeredygtigudvikling.nu &nbsp;&nbsp;&nbsp;&nbsp; <br>
<input type="checkbox" name="econet" value="ok" <%if ((rsPageData.Fields.Item("econetsite").Value)=1) then response.write("checked")%>>
eco-net.dk &nbsp;&nbsp;&nbsp;&nbsp; <br>
Alle nyheder vises p&aring; eco-info.dk<br>
&nbsp; 
<input type="hidden" name="id" value="<%=request("id")%>">
</td>
</tr>
<tr> 
<td width="30%" bgcolor="#CCCCCC"><span class="listheader">&nbsp;Linktekst:</span> 
<input type="text" name="linktext2" value="<%=(rsPageData.Fields.Item("linktext2").Value)%>" class="formInputobject">
</td>
<td width="30%" valign="top" bgcolor="#CCCCCC"><span class="listheader">Link:</span> 
<input value="<%=(rsPageData.Fields.Item("FileName2").Value)%>" type="text" name="link2" size="32" class="formInputobject">
</td>
</tr>
<tr> 
<td width="30%" bgcolor="#CCCCCC"><span class="listheader">&nbsp;Linktekst:</span> 
<input type="text" name="linktext3" value="<%=(rsPageData.Fields.Item("linktext3").Value)%>" class="formInputobject">
</td>
<td width="30%" valign="top" bgcolor="#CCCCCC"><span class="listheader">Link:</span> 
<input value="<%=(rsPageData.Fields.Item("FileName3").Value)%>" type="text" name="link3" size="32" class="formInputobject">
</td>
</tr>
<tr> 
<td width="60%" colspan="2" bgcolor="#CCCCCC"><span class="listheader"> 
<input type="radio" name="vindue" <%if ((rsPageData.Fields.Item("newwin").Value)=0) then response.write("checked")%> value="0">
&Aring;bnes i nyt vindue <br>
<input type="radio" name="vindue" <%if ((rsPageData.Fields.Item("newwin").Value)=1) then response.write("checked")%> value="1">
&Aring;bnes i samme vindue<br>
<br>
</span><span class="formLabeltext"> 
<input type="hidden" name="opdater" value="1">
<input type="hidden" name="sletNyhed">
</span></td>
<td width="40%"> 
<div align="center">Hvis nyheden tilh&oslash;rer et tema:<br>
<select name="tema" class="formInputobject">
<option selected value="0">vælg tema 
<option> 
<%
While (NOT rsTema.EOF)
%>
<option value="<%=(rsTema.Fields.Item("id").Value)%>" <%if rsPageData("Tema_ID")=rsTema("id") then response.write("selected")%>><%=(rsTema.Fields.Item("name").Value)%> </option>
<%
  rsTema.MoveNext()
Wend
If (rsTema.CursorType > 0) Then
  rsTema.MoveFirst
Else
  rsTema.Requery
End If
%>
</select>
</div>
</td>
</tr>
<tr> 
<td width="40%" colspan="2"> 
<input type="hidden" name="createDate" value="<%=CStr(Date)%>">
<span class="listheader">Inds&aelig;t billede i h&oslash;jre kolonne (Rightbar)</span><br>
<a href="/admin/ei/billedbasen/index.asp?size=4" target="_blank"><span class="listheader">Find 
Billede</span></a> Skriv Billednr: 
<input type="text" name="img_id" size="4" value="<%=(rsPageData.Fields.Item("img_id").Value)%>">
</td>
<td width="60%"> 
<div align="left"><span class="formLabeltext"> 
<input type="checkbox" name="validated" value="ok" <%if (rsPageData.Fields.Item("Approved").Value)=1 then response.write("checked")%>>
<span class="formLabeltext">Godkendt</span> <br>
<input type="submit" name="Submit" value="Gem" class="formSubmitbutton" onClick="MM_validateForm('title','','R','shortdescription','','R');OnFormSubmit();return document.MM_returnValue">
<input type="button" name="SubmitSlet" value="Slet" class="formSubmitbutton" onClick="deleteNews()">
</span></div>
</td>
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
<%
rsTema.Close()
%>
<!--#include virtual="/shared/log.asp"-->

