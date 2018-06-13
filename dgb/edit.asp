<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/inc_access.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->


<%
Dim rsOrg__ColParam
rsOrg__ColParam = "0"
if (request("id") <> "") then rsOrg__ColParam = request("id")
%>

<%
set rsOrg = Server.CreateObject("ADODB.Recordset")
rsOrg.ActiveConnection = MM_ecoinfo_STRING
rsOrg.Source = "SELECT o.id,o.name, o.firstname,o.lastname  FROM eiorg_maindata o LEFT JOIN eilib_r_org l ON o.id=l.orgid  WHERE l.libid=" + Replace(rsOrg__ColParam, "'", "''") + ""
rsOrg.CursorType = 0
rsOrg.CursorLocation = 2
rsOrg.LockType = 3
rsOrg.Open()
rsOrg_numRows = 0
%>

<%
Dim rsPageData__ColParam
rsPageData__ColParam = "0"
if (request("id") <> "") then rsPageData__ColParam = request("id")
%>
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT *  FROM eilib_maindata  WHERE id=" + Replace(rsPageData__ColParam, "'", "''") + ""
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>
<%
Dim rsCategories__ColParam
rsCategories__ColParam = "0"
if (request("id") <> "") then rsCategories__ColParam = request("id")
%>
<%
set rsCategories = Server.CreateObject("ADODB.Recordset")
rsCategories.ActiveConnection = MM_ecoinfo_STRING
rsCategories.Source = "SELECT c.id,c.name  FROM eilib_cat_maindata c LEFT JOIN eilib_r_cat o ON c.id=o.catid  WHERE o.libid=" + Replace(rsCategories__ColParam, "'", "''") + ""
rsCategories.CursorType = 0
rsCategories.CursorLocation = 2
rsCategories.LockType = 3
rsCategories.Open()
rsCategories_numRows = 0
%>
<%
Dim rsSubjects__ColParam
rsSubjects__ColParam = "0"
if (request("id") <> "") then rsSubjects__ColParam = request("id")
%>
<%
set rsSubjects = Server.CreateObject("ADODB.Recordset")
rsSubjects.ActiveConnection = MM_ecoinfo_STRING
rsSubjects.Source = "SELECT s.id,s.name  FROM eilib_r_sbj o LEFT JOIN eisbj_maindata s ON o.sbjid=s.id  WHERE o.libid=" + Replace(rsSubjects__ColParam, "'", "''") + ""
rsSubjects.CursorType = 0
rsSubjects.CursorLocation = 2
rsSubjects.LockType = 3
rsSubjects.Open()
rsSubjects_numRows = 0
%>
<%
set rsAllCats = Server.CreateObject("ADODB.Recordset")
rsAllCats.ActiveConnection = MM_ecoinfo_STRING
rsAllCats.Source = "SELECT id,name  FROM eilib_cat_maindata  WHERE siteid=1"
rsAllCats.CursorType = 0
rsAllCats.CursorLocation = 2
rsAllCats.LockType = 3
rsAllCats.Open()
rsAllCats_numRows = 0
%>
<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Ecoinfo</title>
<script src="/shared/multiselect.js"></script>
<script src="lib.js"></script>
<script src="/shared/showeditor.js"></script>
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function AddToLayer(objName,x,newText) { //v4.01
  if ((obj=MM_findObj(objName))!=null) with (obj)
	innerHTML += unescape(newText);
}

function setdgs(theid,thename){
	AddToLayer('dgsnames','','<br>'+thename+'<input type="hidden" name="theorgid2" value="'+theid+'">')
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
curtab=3
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
p&aring;, at felter markeret med en * skal v&aelig;re udfyldt.<br>
<br>
<span class="sidebarHeader">P&aring; nettet</span><br>
Udfyldes hvis publikationen er en web-side, eller hvis en n&aelig;rmere beskrivelse, 
bestilling o.l. kan foretages der.</p>
<p><span class="sidebarHeader">Om udgiveren</span><br>
Disse felter udfyldes kun, hvis du/I ikke selv har udgivet materialet.</p>
<p><span class="sidebarHeader">Kort og uddybende beskrivelse</span><br>
Den korte beskrivelse vises i oversigten, mens b&aring;de kort og uddybende vises, 
n&aring;r nogen ser p&aring; dit kort i De Gr&oslash;nne Sider.</p>
<p><span class="sidebarHeader">Forsk&oslash;n teksten med HTML-formatering.</span> 
<br>
<a href="#" onClick="MM_openBrWindow('/shared/HTMLhelp.htm','HTMLhj&aelig;lp','scrollbars=yes,width=800,height=600')">For 
MAC brugere.</a></p>
<p><span class="sidebarHeader">Kategori og emneord</span><br>
- er vigtige for at du bliver vist rette sted. Dine valg her er kun vejledende 
for vores redakt&oslash;rer der foretager den endelige kategorisering.</p>
<p><span class="sidebarHeader">Stikord</span><br>
P&aring; &Oslash;ko-info kan man lave fritekst-s&oslash;gninger. Hvis du &oslash;nsker 
at blive vist n&aring;r der s&oslash;ges p&aring; ord der ikke findes i dine beskrivelser, 
kan de skrives her.<br>
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
<form method="post" action="act_edit.asp" onSubmit="doedit(document.forms[0]);return document.validation">
<tr> 
<td valign="top"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
<tr> 
<td class="contentHeader1"> Rediger publikation 
<%IF LEN(request.cookies("eiuserid"))>0 THEN%>
&nbsp;(id: <%=rsPageData("id")%>) 
<%end if %>
</td>
</tr>
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top"> 
<% If Not rsOrg.EOF Or Not rsOrg.BOF Then %>
<%IF LEN(request.cookies("eiuserid"))>0 THEN
	response.write "Tilknyttet: "
	IF LEN(rsOrg.Fields.Item("name").value)>0 THEN
		response.write rsOrg.Fields.Item("name").value & "<br><br>"
	ELSE
		response.write rsOrg.Fields.Item("firstname").value & " " & rsOrg.Fields.Item("lastname").value & "<br><br>"
	END IF
END IF%>
<br>
<% End If ' end Not rsOrg.EOF Or NOT rsOrg.BOF %>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr valign="top"> 
<td> Om Publikationen<br>
<span class="formLabeltext">*Titel:</span><br>
<input type="text" name="title" value="<%=(rsPageData.Fields.Item("title").Value)%>" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext">Undertitel</span>: <br>
<input type="text" name="subtitle" value="<%=(rsPageData.Fields.Item("subtitle").Value)%>" size="32" class="formInputobjectLong">
<br>
<%IF LEN(request.cookies("eiuserid"))>0 THEN%>
<span class="formLabeltext">*Kategori:</span><br>
<select name="selCat" class="formInputobjectLong">
<option value="">&lt; --- &gt;</option>
<%
IF NOT rsCategories.EOF THEN 
	cursel=CStr(rsCategories.Fields.Item("id").Value & "")
ELSE 
	cursel=""%>
<%END IF
While (NOT rsAllCats.EOF)
%>
<option value="<%=(rsAllCats.Fields.Item("id").Value)%>" <%if (CStr(rsAllCats.Fields.Item("id").Value & "") = cursel) then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsAllCats.Fields.Item("name").Value)%></option>
<%
  rsAllCats.MoveNext()
Wend
If (rsAllCats.CursorType > 0) Then
  rsAllCats.MoveFirst
Else
  rsAllCats.Requery
End If
%>
</select>
<br>
<input name="islocal" type="checkbox" value="1" <%if CINT("0"&rsPageData("islocal"))=1 then response.write "checked"%>> Øko-net har den<br><br>
<%END IF%>
<span class="formLabeltext">*Forfatter:</span><br>
<input type="text" name="author" value="<%=(rsPageData.Fields.Item("author").Value)%>" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext">ISBN:</span><br>
<input type="text" name="isbn" value="<%=(rsPageData.Fields.Item("isbn").Value)%>" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext">Udgivelsesår:</span><br>
<input type="text" name="year" value="<%=(rsPageData.Fields.Item("year").Value)%>" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext">Sprog:</span><br>
<input type="text" name="language" value="<%=(rsPageData.Fields.Item("language").Value)%>" size="32" class="formInputobjectLong">
<br>
<span class="formLabeltext">P&aring; nettet:</span><br>
<input type="text" name="webaddress" value="<%=(rsPageData.Fields.Item("webaddress").Value)%>" size="32" class="formInputobjectLong">
</td>
<td> <p>Om udgiveren<br>
      <span class="formLabeltext">Udgiver/Forlag:</span><br>
      <textarea name="editor" cols="32" class="formTextobjectLow"><%=(rsPageData.Fields.Item("editor").Value)%></textarea>
      <span class="formLabeltext"><br>
  Email-adresse:</span><br>
  <input type="text" name="editoremail" value="<%=(rsPageData.Fields.Item("editoremail").Value)%>" size="32" class="formInputobjectLong">
  <span class="formLabeltext"><br>
  P&aring; nettet:</span><br>
  <input type="text" name="editorwww" value="<%=(rsPageData.Fields.Item("editorwww").Value)%>" size="32" class="formInputobjectLong">
</p>
  <p align="left">
    <input type="submit" name="Submit2" value="Gem" class="formSubmitbutton">
  </p></td>
</tr>
<%IF LEN(request.cookies("eiuserid"))>0 THEN %>
<tr><td colspan="2">
<b>I De Gr&oslash;nne Sider:</b><br>
<input type="button" name="Button" value="V&aelig;lg" class="formbutton" onClick="window.open('/shared/sel_org.asp','subwin','scrollbars=yes,resizable=yes,width=500,height=500,top=100,left=200');">
<div id="dgsnames"> 
<% If Not rsOrg.EOF Or Not rsOrg.BOF Then %>
<input type="hidden" name="firstorgid" value="<%=rsOrg.Fields.Item("id").value%>">
<table border="0" cellpadding="0" cellspacing="0">
<%
i=0
WHILE NOT rsOrg.EOF
	IF LEN(rsOrg.Fields.Item("name").value)>0 THEN
		thename=rsOrg.Fields.Item("name").value
	ELSE
		thename=rsOrg.Fields.Item("firstname").value & " " & rsOrg.Fields.Item("lastname").value
	END IF
	response.write "<tr><td>" & thename & "</td><td><input type=""hidden"" name=""theorgid2"" value=""" & rsOrg("id") & """><input type=""checkbox"" name=""delorg"" value=""" & rsOrg("id") & """ onclick=""this.form.theorgid2[" & i & "].value='';"">Slet</td></tr>"
	rsOrg.movenext
	i=i+1
WEND
%>
</table>
<%End If ' end Not rsOrg.EOF Or NOT rsOrg.BOF %>
</div>
<input type="hidden" value="" name="theorgid">

</td></tr>
<%END IF %>
</table>
<br>
Beskrivelser 
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr valign="top"> 
<td width="50%"> <span class="formLabeltext">*Kort beskrivelse:</span><br>
<textarea name="shortdescription" cols="50" rows="5" class="formTextobjectLow"><%=(rsPageData.Fields.Item("shortdescription").Value)%></textarea>
<div align="center"></div>
</td>
<td> <span class="formLabeltext">*Resum&eacute;:</span><br>
<textarea name="description" cols="50" rows="5" class="formTextobjectBig"><%=(rsPageData.Fields.Item("description").Value)%></textarea>
<div align="center"> 
<input type="button" class="function" onClick="showeditor('description','');" value="Formater" name="button">
</div>
</td>
</tr>
</table>
<%IF LEN(request.cookies("eiuserid"))>0 THEN%>
*Emneord 
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr valign="top"> 
<td width="50%"> <span class="formLabeltext">Alle emneord:</span><br>
<select name="allSbj" class="formTextobjectLow" size="5"  ondblclick="addValue(this.form.allSbj,this.form.selSbj);">
<script src="/log/insider/ei/menufiles/sbj_options.js"></script>
</select>
<br>
Dobbeltklik p&aring; et emneord for at v&aelig;lge det </td>
<td> <span class="formLabeltext">Valgte Emneord:</span><br>
<select name="selSbj" class="formTextobjectLow" size="5" ondblclick="removeValue(this.form.selSbj);" multiple>
<%
While (NOT rsSubjects.EOF)
%>
<option value="<%=(rsSubjects.Fields.Item("id").Value)%>"><%=(rsSubjects.Fields.Item("name").Value)%></option>
<%
  rsSubjects.MoveNext()
Wend
If (rsSubjects.CursorType > 0) Then
  rsSubjects.MoveFirst
Else
  rsSubjects.Requery
End If
%>
</select>
<br>
Dobbeltklik p&aring; et emneord for at fjerne det </td>
</tr>
</table>
<%ELSE %>
<br>
<span class="listheader"> Valgte Emneord:</span> 
<%rsSubjects.MoveFirst()
While (NOT rsSubjects.EOF)
  response.write (rsSubjects.Fields.Item("name").Value) & ", "
  rsSubjects.MoveNext()
Wend
%>
<br>
Ønsker du at ændre emneord, <a href="javascript:MM_openBrWindow('user_catsbj.asp?id=<%=request("id")%>&name=<%=rsPageData("title")%>','Emneord','width=600,height=600')">klik 
her</a> 
<%END IF%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr valign="top"> 
<td width="50%"> <span class="formLabeltext"><br>
Stikord:</span><br>
<textarea name="keywords" class="formTextobjectLow"><%=(rsPageData.Fields.Item("keywords").Value)%></textarea></td>
<td><p>&nbsp;</p>
  <p>&nbsp;</p></td>
</tr>

<tr valign="top">
  <td><p>&nbsp;</p>
    <p><strong>Billeder: </strong></p>
    <%IF LEN(request.cookies("eiuserid"))>0 THEN%>
        <span class="listheader">Inds&aelig;t billede under tekst <br>
billedet skal findes som 3 col (3)</span><br>
<a href="/admin/ei/billedbasen/index.asp?sizes=5" target="_blank"><span class="listheader">Find 
    Billede</span></a> Skriv Billednr:
    <input value="<%=(rsPageData.Fields.Item("img_id").Value)%>" type="text" name="img_id" size="4">
</p>
    <p><a href="/shared/picasa.asp" target="_blank"><strong>Billede fra Picasa </strong></a><br>
        <input name="imagefilename1" type="text" id="imagefilename1" value="<%=rsPageData("imagefilename1")%>" size="40">
        <% end if %>
    </p></td>
  <td>
    <p>&nbsp;</p>
    <p>&nbsp; </p>
    <p>
      <%IF LEN(request.cookies("eiuserid"))>0 THEN%>
      <span class="listheader">Inds&aelig;t billede i h&oslash;jre side<br>
        billedet skal findes som rightcol (R)</span><br>
  <a href="/admin/ei/billedbasen/index.asp?sizes=4" target="_blank"><span class="listheader">Find 
    Billede</span></a> Skriv Billednr:
      <input value="<%=(rsPageData.Fields.Item("img_id2").Value)%>" type="text" name="img_id2" size="4">
    </p>
    <p><a href="/shared/picasa.asp" target="_blank"><strong>Billede fra Picasa</strong></a><br>
        <input name="imagefilename2" type="text" id="imagefilename2" value="<%=rsPageData("imagefilename2")%>" size="40">
        <% end if %>
    </p></td>
</tr>
</table>
<div align="center"> 
<%IF LEN(request.cookies("eiuserid"))=0 THEN%>
<input type="hidden" name="orgid" value="<%=request.cookies("eiorgid")%>">
<%END IF%>
<!--#include virtual="/shared/inc_getparams.asp" -->
<%=GetParams("&id=")%> 
<input type="hidden" name="id" value="<%=(rsPageData.Fields.Item("id").Value)%>">
<br>
<input type="reset" value="Fortryd &aelig;ndringer" class="formInputobject" name="Reset">
<input type="submit" name="Submit" value="Gem" class="formSubmitbutton">
<!--#include virtual="/shared/inc_validate.asp"-->
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
<%
rsPageData.Close()
%>
<%
rsCategories.Close()
%>
<%
rsSubjects.Close()
%>
<%
rsAllCats.Close()
%>
<%
rsOrg.Close()
%>
<!--#include virtual="/shared/log.asp"-->






