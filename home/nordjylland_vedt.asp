<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT *  FROM bestyrelse_maindata  ORDER BY orden"
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsPageData_numRows = rsPageData_numRows + Repeat1__numRows
%>
<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Ecoinfo</title>
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
<!-- #BeginEditable "menu" --> <!-- #BeginLibraryItem "/Library/menu.lbi" --><table width="750" border="0" cellspacing="0" cellpadding="0" name="Menu">
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
<%level1=6%>
<!--#include file="inc_about_leftbar.asp"-->
<!-- #EndEditable --></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="569" valign="top" height="100%" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> <!-- #BeginEditable "maincontent" -->
  <table width="100%" border="0" cellspacing="0" cellpadding="16" background="/dgs/graphics/detail_header_dgs_bkgrd.gif" name="detailHeader">
    <tr>
      <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif" name="detailContents">
          <tr>
            <td colspan="2" class="contentHeader1" align="left">Vedt&aelig;gter for &Oslash;ko-net Nordjylland </td>
          </tr>
          <tr>
            <td colspan="2" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
          </tr>
          <tr>
            <td height"20">&nbsp;</td>
            <td height="20" align="right" nowrap>&nbsp;</td>
          </tr>
        </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="/shared/graphics/spacer.gif" width="1" height="5"></td>
            </tr>
          </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td valign="top"><table border="0" cellspacing="0" cellpadding="0" width="100%">
                  <tr>
                    <td><p><strong>&sect; 1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Navn og hjemsted</strong></p>
                        <p>&Oslash;ko-net Nordjylland er en lokalafdeling tilknyttet Netv&aelig;rket  for &oslash;kologisk folkeoplysning og praksis/&Oslash;ko-net, der er landsforeningen og er  hjemmeh&oslash;rende i Svendborg Kommune i Region Syddanmark.<br>
                          &Oslash;ko-net Nordjylland er hjemmeh&oslash;rende i &Aring;lborg Kommune i  Region Nordjylland.</p>
                      <p><strong>&sect; 2&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Form&aring;l</strong></p>
                      <p>Form&aring;let med &Oslash;ko-net Nordjylland er p&aring; et folkeligt niveau  at informere, oplyse og inspirere omkring natur og milj&oslash;, &oslash;kologi og b&aelig;redygtig  udvikling - samt at skabe debat og netv&aelig;rk omkring &oslash;kologiske og b&aelig;redygtige  tiltag.<br>
                        Ved &oslash;kologi forst&aring;s en husholdning med ressourcer, der er i  balance med naturen.<br>
                        Ved b&aelig;redygtig udvikling forst&aring;s en udvikling, der tager  globale, milj&oslash;m&aelig;ssige og sociale hensyn b&aring;de til nulevende og kommende  generationer.<br>
                        Der l&aelig;gges v&aelig;gt p&aring; ogs&aring; at engagere ungdommen i foreningen  og i arbejdet med ovenn&aelig;vnte form&aring;l.</p>
                      <p><strong>&sect; 3&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Medlemmer af foreningen</strong></p>
                      <p>Medlemmer af &Oslash;ko-net Nordjylland er alle som kan tilslutte  sig &sect;2 og betaler det af landsforeningen fastsatte landskontingent p&aring; &Oslash;ko-nets  &aring;rlige generalforsamling. Medlemskab af &Oslash;ko-net Nordjylland kr&aelig;ver derfor at  man er medlem af landsorganisationen &Oslash;ko-net. Der opkr&aelig;ves kun kontingent via  landsforeningen.</p>
                      <p><strong>&sect; 4&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Generalforsamling</strong></p>
                      <p>&Oslash;ko-net Nordjyllands h&oslash;jeste myndighed er  generalforsamlingen, der ordin&aelig;rt afholdes hvert &aring;r i for&aring;ret. Indkaldelse til  generalforsamlingen sker gennem medierne eller ved underretning af medlemmerne  p&aring; anden m&aring;de, mindst 1 m&aring;ned forud. Forslag, som man &oslash;nsker dr&oslash;ftet p&aring;  generalforsamlingen, skal v&aelig;re bestyrelsen i h&aelig;nde 14 dage forud for m&oslash;det.</p>
                      <p>Dagsorden for ordin&aelig;r generalforsamling skal indeholde  f&oslash;lgende punkter:<br>
                        1. Valg af dirigent og referent.<br>
                        2. Bestyrelsens beretning.<br>
                        3. Freml&aelig;ggelse af &aring;rsregnskab.<br>
                        4. Indkomne forslag.<br>
                        5. Valg af bestyrelsesmedlemmer, samt suppleant(er)<br>
                        6. Valg af revisor.<br>
                        7. Eventuelt</p>
                      <p><strong>&sect; 6&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Bestyrelsen</strong></p>
                      <p>Stk. 1 Generalforsamlingen v&aelig;lger en bestyrelse. Der v&aelig;lges  indtil 5 medlemmer og 2 suppleanter. Efter valg konstituerer bestyrelsen sig  med formand og n&aelig;stformand. I tilf&aelig;lde af formandens frav&aelig;r fungerer  n&aelig;stformanden som formand. </p>
                      <p>Stk. 2 Bestyrelsesvalg g&aelig;lder for 2 &aring;r. Halvdelen afg&aring;r  hvert &aring;r. Genvalg kan finde sted.</p>
                      <p>Stk. 3 Bestyrelsen er beslutningsdygtig, n&aring;r mindst  halvdelen af medlemmerne er til stede. I tilf&aelig;lde af stemmelighed g&oslash;r  formandens og i dennes frav&aelig;r n&aelig;stformandens stemme udslaget.</p>
                      <p>Stk. 4 Bestyrelsen kan fasts&aelig;tte en forretningsorden for  bestyrelsens arbejde.</p>
                      <p>Stk. 5 Bestyrelsen f&oslash;rer referat over sine m&oslash;der, der p&aring; det  efterf&oslash;lgende m&oslash;de godkendes.</p>
                      <p><strong>&sect; 7&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &Oslash;konomi</strong></p>
                      <p>Stk. 1 Lokalafdelingens midler kan tilvejebringes gennem  lokale folkeoplysnings- og fondsmidler og projektst&oslash;tte fra landsforeningen,  n&aring;r og hvis dette er muligt.</p>
                      <p>Stk. 2 &Oslash;ko-net Nordjyllands regnskabs&aring;r f&oslash;lger kalender&aring;ret.</p>
                      <p>Stk. 3 &Oslash;ko-net Nordjyllands &aring;rsregnskab forel&aelig;gges og  kommenteres af den valgte revisor.</p>
                      <p><strong>&sect; 8&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Ekstraordin&aelig;r  generalforsamling</strong></p>
                      <p>Ekstraordin&aelig;r generalforsamling indkaldes n&aring;r bestyrelsen  eller mindst en 1/3 af medlemmerne i lokalafdelingen beg&aelig;rer det.</p>
                      <p><strong>&sect; 9&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Vedt&aelig;gts&aelig;ndringer og  nedl&aelig;ggelse</strong></p>
                      <p>Stk. 1 Vedt&aelig;gterne kan &aelig;ndres p&aring; en generalforsamling, n&aring;r  forslaget er udsendt mindst 2 uger f&oslash;r generalforsamlingen og n&aring;r mindst 2/3 af  de fremm&oslash;dte stemmer derfor. Den enkelte lokalafdeling kan tilf&oslash;je punkter til  disse vedt&aelig;gter, s&aring;fremt der er behov for det. Dog m&aring; punkterne ikke stride mod  n&aelig;rv&aelig;rende vedt&aelig;gter eller &Oslash;ko-nets landsvedt&aelig;gter. N&aelig;rv&aelig;rende vedt&aelig;gter og  eventuelle senere &aelig;ndringer skal godkendes af &Oslash;ko-nets landforening.</p>
                      <p>Stk. 2 Opl&oslash;sning af &Oslash;ko-net Nordjylland kan ske ved  beslutning af s&aring;vel den ordin&aelig;re som af en ekstraordin&aelig;rt indkaldt  generalforsamling. Vedtagelse af beslutning om &Oslash;ko-net Nordjylland opl&oslash;sning  kr&aelig;ver, at mindst tre fjerdedele af medlemmerne er repr&aelig;senterede ved  generalforsamlingen og at to tredjedele af disse stemmer for opl&oslash;sningen af  lokal afdelingen.<br>
                        Hvis ikke tre fjerdedele af &Oslash;ko-net Nordjyllands medlemmer  er repr&aelig;senteret ved generalforsamlingen, indkalder bestyrelsen til en ny  ekstraordin&aelig;r generalforsamling, hvor opl&oslash;sningen af lokalafdelingen kan  vedtages af to tredjedele af de repr&aelig;senterede.</p>
                      <p>Stk. 3 S&aring;fremt opl&oslash;sningen vedtages og der ved likvidationen  er midler i behold ud over, hvad der er forn&oslash;dent til d&aelig;kning af &Oslash;ko-net  Nordjylland forpligtelser, g&aring;r midlerne til &Oslash;ko-nets landsforening.</p>
                      <p>S&aring;ledes vedtaget p&aring;  den stiftende generalforsamling for &Oslash;ko-net Nordjylland den 19/12-2006</p></td>
                  </tr>
                  <tr>
                    <td valign="top"><p>&nbsp;</p>
                        <p><br>
                      </p></td>
                  </tr>
              </table></td>
              <td width="8"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
              <td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
              </td>
              <td width="8"><img src="/shared/graphics/spacer.gif" width="8" height="1"></td>
              <td valign="top" width="200"><table border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="188" class="contentHeader2"><strong>Kontakt information </strong>: </td>
                  </tr>
                  <tr>
                    <td><a href="hovedstaden.asp">Hovedstaden</a><br>
                        <a href="sjaelland.asp">Sj&aelig;lland</a><br>
                        <a href="midtjylland.asp">Midtjylland</a><br>
                        <a href="nordjylland.asp">Nordjylland</a><br>
                        <br></td>
                  </tr>
                  <tr>
                    <td><strong>Vedt&aelig;gter:</strong></td>
                  </tr>
                  <tr>
                    <td><a href="hovedstaden_vedt.asp">Hovedstaden</a><br>
                        <a href="sjaelland_vedt.asp">Sj&aelig;lland</a><br>
                        <a href="midtjylland_vedt.asp">Midtjylland</a><br>
                        <a href="nordjylland_vedt.asp">Nordjylland</a><br>
                        <br></td>
                  </tr>
                  <tr>
                    <td><img src="/shared/graphics/Nordjylland.gif" width="200"></td>
                  </tr>
              </table></td>
            </tr>
        </table></td>
    </tr>
  </table>
  <br>
  <br>
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
<!--#include virtual="/shared/log.asp"-->

