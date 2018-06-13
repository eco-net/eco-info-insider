<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/inc_access.asp" -->

<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include file="act_sqlcalc.asp" -->
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT m.id,m.title,m.shortdescription,m.www,m.startdate,m.enddate,m.status,m.modidate,p.city,m.maker,m.editor,m.createdate  FROM (eical_maindata m LEFT JOIN eisys_postnr p ON m.postnr=p.postnr)  WHERE 0=0  ORDER BY m.startdate"
rsPageData.Source=replace(rsPageData.Source,"0=0",theWhere)
rsPageData.Source=replace(rsPageData.Source,"(eical_maindata m LEFT JOIN eisys_postnr p ON m.postnr=p.postnr)",theFrom)
rsPageData.Source=replace(rsPageData.Source,"ORDER BY m.startdate",theOrder)
'response.write rsPageData.Source
'response.end
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 10
Dim Repeat1__index
Repeat1__index = 0
rsPageData_numRows = rsPageData_numRows + Repeat1__numRows
%>
<!--#include virtual="/shared/inc_listnavigation.asp"-->
<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Øko-info Insider</title>
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
curtab=2%>
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
<%IF LEN(request.cookies("eiorgid"))=0 THEN%>
<!--#include file="inc_search.asp"-->
<%ELSE%>
<!--#include file="inc_info.asp"-->
<!--#include virtual="/shared/inc_login.asp"-->
<%END IF%>
<!-- #EndEditable --></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="569" valign="top" height="100%" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> <!-- #BeginEditable "maincontent" --> 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top" width="98%"> 
<% If Not rsPageData.EOF Or Not rsPageData.BOF Then %>
<%IF LEN(request.cookies("eiuserid"))>0 THEN%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="8" height="30"></td>
<td class="contentHeader1" colspan="2" width="98%">Fandt <%=rsPageData_total%> arrangementer <font color="#FF0000"><%=validtext%></font></td>
<td><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
</tr>
<tr> 
<td colspan="4" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr valign="top"> 
<td><img src="/shared/graphics/spacer.gif" width="8" height="30"></td>
<td><%=recInfo%></td>
<td align="right" nowrap><%=naviHtml%></td>
<td><br>
</td>
</tr>
</table>
<%END IF%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td colspan="3"><img src="/shared/graphics/spacer.gif" width="1" height="1"></td>
</tr>
<tr> 
<td><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
<td valign="top" width="99%"> <br>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<%
IF LEN(request.cookies("eiuserid"))=0 THEN Repeat1__numRows=rsPageData_total
While ((Repeat1__numRows <> 0) AND (NOT rsPageData.EOF)) 

response.write "<tr><td colspan='2'><img src=""/shared/graphics/spacer.gif"" width=""3"" height=""4""></td></tr>"
response.write "<tr><td><span class=""contentheader2""><a href=""" & replace(detailLink,"##",rsPageData("id")) & """  title=""Viser alle oplysninger om dette arrangement"">"
response.write rsPageData.Fields.Item("title").Value
response.write "</a></span>&nbsp;&nbsp;&nbsp;"
IF LEN(request.cookies("eiuserid"))>0 then
response.write "<a href=""" & replace(replace(detailLink,"##",rsPageData("id")),"edit","delete") & """>Slet</a>" 
response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href=""copy.asp?calid=" & rsPageData("id") & """>" & "Kopiér</a>"
end if
response.write("<br>")
IF DatePart("m",rsPageData.Fields.Item("startdate").value)=DatePart("m",rsPageData.Fields.Item("enddate").value) THEN
	IF DatePart("d",rsPageData.Fields.Item("startdate").value)=DatePart("d",rsPageData.Fields.Item("enddate").value) THEN
		response.write FormatDateTime(rsPageData.Fields.Item("startdate").value,vbLongDate)
	ELSE
		response.write DatePart("d",rsPageData.Fields.Item("startdate").value) & " til " & FormatDateTime(rsPageData.Fields.Item("enddate").value,vbLongDate)
	END IF
ELSEIF rsPageData.Fields.Item("enddate").value<>"" THEN
	response.write FormatDateTime(rsPageData.Fields.Item("startdate").value,vbLongDate) & " til " & FormatDateTime(rsPageData.Fields.Item("enddate").value,vbLongDate)
ELSE
	response.write FormatDateTime(rsPageData.Fields.Item("startdate").value,vbLongDate)
END IF
'IF rsPageData.Fields.Item("city").value<>"" THEN response.write " | " & rsPageData.Fields.Item("city").value
'response.write "<br>" & replace(rsPageData.Fields.Item("shortdescription").Value & "",chr(13),"<br>") & "</td></tr>"

response.write "</td><td align='right'><font color='red'>"
if request("valid")<>"" then
if rsPageData.Fields.Item("maker")<>"" or rsPageData.Fields.Item("editor")<>"" then
response.write ("Oprettet: " & DateValue(rsPageData.Fields.Item("createdate")) & " af " & rsPageData.Fields.Item("maker") & "</font><br>")
if rsPageData.Fields.Item("editor")<>"" then 
response.write ("<font color='red'>Rettet: " & DateValue(rsPageData.Fields.Item("modidate")) & " af " & rsPageData.Fields.Item("editor") & "</font>")
end if
else 'maker/editor
if len(rsPageData.Fields.Item("status"))>0 then
	if rsPageData.Fields.Item("status")=1 then 
	response.write "&nbsp;&nbsp;Oprettet af Øko-net "
	elseif rsPageData.Fields.Item("status")=2 then 
	response.write "&nbsp;&nbsp;Oprettet af Bruger "
	elseif rsPageData.Fields.Item("status")=3 then 
	response.write "&nbsp;&nbsp;Rettet af Øko-net "
	elseif rsPageData.Fields.Item("status")=4 then 
	response.write "&nbsp;&nbsp;Rettet af Bruger "
	end if
	response.write FormatDateTime(rsPageData("modidate"),vbShortDate)
end if
end if
end if
response.write "</font></td></tr>"

response.write "<tr><td colspan='2'><img src=""/shared/graphics/spacer.gif"" width=""3"" height=""7""></td></tr>"
response.write "<tr><td colspan='2' background=""/shared/graphics/layout/dots_horz.gif"" height=""1""><img src=""/shared/graphics/spacer.gif"" width=""3"" height=""1""></td></tr>"

  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsPageData.MoveNext()
Wend
%>
</table>
<br>
<br>
</td>
<td><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
</tr>
<tr> 
<td><br>
</td>
<td valign="top" align="right"> 
<%IF LEN(request.cookies("eiuserid"))>0 THEN response.write naviHtml%>
</td>
<td><br>
</td>
</tr>
</table>
<% End If ' end Not rsPageData.EOF Or NOT rsPageData.BOF %>
<% If rsPageData.EOF And rsPageData.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td width="8"><br>
</td>
<td class="contentHeader1" height="22">Fandt ingen arrangementer</td>
<td width="8"><br>
</td>
</tr>
<tr> 
<td colspan="3" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td><br>
</td>
<td> <br>
<br>
<br>
Der er ingen arrangementer i &Oslash;ko-kalenderen, der matcher din søgning.<br>
Det skyldes sansynligvis, at du har lavet en meget pr&aelig;cis s&oslash;gning. 
</td>
<td><br>
</td>
</tr>
</table>
<% End If ' end rsPageData.EOF And rsPageData.BOF %>
</td>
<td valign="top" height="100%"> </td>
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
<%'reg("oklist")%>