<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/common.asp" -->
<%
Dim rsPageData__ColParam
rsPageData__ColParam = "0"
if (request.cookies("eiorgid") <> "") then rsPageData__ColParam = request.cookies("eiorgid")
%>

<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT * FROM eiorg_maindata WHERE id=" + Replace(rsPageData__ColParam, "'", "''") + ""
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>


<!--#include file="inc_redirect.asp"-->
<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Øko-net Insider</title>
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
curtab=1%>
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
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;

OM STAMDATA</td>
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

Her kan du redigere oplysninger i De Gr&oslash;nne Sider, samt indstille brugen af data fra &Oslash;ko-info p&aring; egen site.<br>
<%IF LEN(request.cookies("eiorgid"))>0 THEN%>
<form name="form2" method="post" action="edit.asp?id=<%=request.cookies("eiorgid")%>">
<div align="center">
<input type="submit" name="Submit" value="Rediger data" class="formInputobject" onClick="this.form.action='edit.asp?id=<%=request.cookies("eiorgid")%>';"><br>
<input type="button" name="Button3" value="Rediger Brugeroplysninger" class="formInputobject" onClick="window.open('users.asp','user','width=400,height=250,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,top=200,left=200')">
<input type="submit" name="Submit" value="Indstillinger" class="formInputobject" onClick="this.form.action='setup.asp?id=<%=request.cookies("eiorgid")%>';"><br>
</div>
</form>
<%END IF%>
</td>
</tr>
<tr>
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</table>
</td>
</tr>
</table>
<table width="180" border="0" cellspacing="0" cellpadding="0">
<tr>
<td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
<td width="98%" class="sidebarHeader">&nbsp;&nbsp;OM INDSTILLINGER</td>
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
<td> Som bruger af &Oslash;ko-info, kan du v&aelig;lge at vise 'dine' data p&aring; 
din egen hjemmeside.<br>
Her bestemmer du, hvorvidt du &oslash;nsker at bruge denne mulighed, samt hvordan 
dataene pr&aelig;senteres.<br>
<br>
</td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="3" height="8"></td>
</tr>
</table>
</td>
</tr>
</table>
<!--#include virtual="/shared/inc_login.asp"-->
 
<!-- #EndEditable --></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="569" valign="top" height="100%" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> <!-- #BeginEditable "maincontent" -->
<table width="100%" border="0" cellspacing="0" cellpadding="16" background="/dgs/graphics/detail_header_dgs_bkgrd.gif" name="detailHeader">
<tr> 
<td valign="top">

<% If Not rsPageData.EOF Or Not rsPageData.BOF Then %>
<span class="contentHeader1">S&aring;dan pr&aelig;senteres dine/Jeres data i De 
Gr&oslash;nne Sider.</span><br>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif" name="detailContents">
<tr> 
<td colspan="2" class="contentHeader1" align="left">
<%IF LEN(rsPageData.Fields.Item("name").Value)=0 THEN
	response.write rsPageData.Fields.Item("firstname").Value & "&nbsp;" & rsPageData.Fields.Item("middlename").Value & "&nbsp;" & rsPageData.Fields.Item("lastname").Value
ELSE
	response.write rsPageData.Fields.Item("name").Value
END IF%>
</td>
</tr>
<tr> 
<td colspan="2" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td height"20"><%=recInfo%></td>
<td height="20" align="right" nowrap><%=naviHTML%></td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td><img src="/shared/graphics/spacer.gif" width="1" height="5"></td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top">
<table border="0" cellspacing="0" cellpadding="0" width="100%">
<tr>
<td class="contentHeader2">Beskrivelse</td>
</tr>
<tr>
<td valign="top">
<p><%=replace(rsPageData.Fields.Item("shortdescription").Value & "",chr(13),"<br>")%></p>
<p><%=replace(rsPageData.Fields.Item("description").Value & "",chr(13),"<br>")%></p>
</td>
</tr>
</table>
</td>
<td width="8"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="8"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
<td valign="top"> 
<table border="0" cellspacing="0" cellpadding="0">
<tr>
<td class="contentHeader2">Fakta</td>
</tr>
<tr> 
<td><span class="faktaboksheader"><br>
Adresse<br>
</span> 
<%IF LEN(rsPageData.Fields.Item("name").Value)=0 THEN
	response.write rsPageData.Fields.Item("firstname").Value & "&nbsp;" & rsPageData.Fields.Item("lastname").Value
ELSE
	response.write rsPageData.Fields.Item("name").Value
END IF%>
<br>
<% if rsPageData.Fields.Item("adrco").Value <>"" then %>
c/o: <%=rsPageData.Fields.Item("adrco").Value%><br>
<%END IF%>
<%if rsPageData.Fields.Item("street1").Value <>"" then%>
<%=replace(rsPageData.Fields.Item("street1").Value,chr(13),"<br>")%> 
<%end if%>
<%if rsPageData.Fields.Item("street2").Value<>"" then%>
<nobr><%=(rsPageData.Fields.Item("street2").Value)%><br>
</nobr> 
<%end if %>
<%if rsPageData.Fields.Item("zip").Value>1 then %>
<nobr><%=(rsPageData.Fields.Item("zip").Value)%> &nbsp;<%=(rsPageData.Fields.Item("city").Value)%></nobr><br>
<% end if %>
<% if rsPageData.Fields.Item("country").Value <> "Danmark" then %>
<%=(rsPageData.Fields.Item("country").Value)%> <br>
<%end if %>
<%if rsPageData.Fields.Item("firstname").Value <>"" or rsPageData.Fields.Item("middlename").Value <>"" or rsPageData.Fields.Item("lastname").Value <>"" then %>
<br>
<span class="listheader">Kontaktperson</span><br>
<%if rsPageData.Fields.Item("title").Value <>"" then%>
<%=(rsPageData.Fields.Item("title").Value)%><br>
<%end if%>
<%=(rsPageData.Fields.Item("firstname").Value)%>&nbsp;<%=(rsPageData.Fields.Item("middlename").Value)%>&nbsp;<%=(rsPageData.Fields.Item("lastname").Value)%> <br>
<%end if%>
<%if rsPageData.Fields.Item("phone1").Value <>"" then%>
<br>
<span class="faktaboksheader">Telefon</span> <br>
<%=splitPhone((rsPageData.Fields.Item("phone1").Value))%> 
<%end if%>
<%if rsPageData.Fields.Item("phone2").Value <>"" then%>
<br>
<%=splitPhone((rsPageData.Fields.Item("phone2").Value))%> 
<%end if%>
<%if rsPageData.Fields.Item("fax").Value <>"" then%>
<br>
<span class="faktaboksheader"><br>
Fax</span><br>
<%=splitPhone((rsPageData.Fields.Item("fax").Value))%> 
<%end if%>
<% if rsPageData.Fields.Item("emailaddress1").Value<>"" or  rsPageData.Fields.Item("emailaddress2").Value<>"" then%>
<span class="faktaboksheader"><br>
<br>
E-mail</span><br>
<a href="mailto:<%=(rsPageData.Fields.Item("emailaddress1").Value)%>" title="Åbner en ny mail i din mail-klient med denne adresse som modtager"><%=(rsPageData.Fields.Item("emailaddress1").Value)%></a> 
<% if rsPageData.Fields.Item("emailaddress2").Value<>"" then %>
<br>
<a href="mailto:<%=(rsPageData.Fields.Item("emailaddress2").Value)%>" title="Åbner en ny mail i din mail-klient med denne adresse som modtager"></a> 
<%end if %>
<%end if %>
<%if rsPageData.Fields.Item("www").Value <>"" then%>
<br>
<span class="faktaboksheader"><br>
Hjemmeside</span><br>
<a href="http://<%=(rsPageData.Fields.Item("www").Value)%>" title="Åbner hjemmesiden i et nyt vindue" target="_blank"><%=rsPageData.Fields.Item("www").Value%><img src="/shared/graphics/layout/arrows_fwd.gif" width="10" height="7" border="0"></a> 
<%end if%>
</td>
</tr>
</table>
</td>
</tr>
</table>
<% End If ' end Not rsPageData.EOF Or NOT rsPageData.BOF %>

<% If rsPageData.EOF And rsPageData.BOF Then %>
<br>
Log ind for at bruge &Oslash;ko-info Insider.

<% End If ' end rsPageData.EOF And rsPageData.BOF %>
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

<!--include virtual="/shared/log.asp"-->
<%'reg("dgsindex")%>

