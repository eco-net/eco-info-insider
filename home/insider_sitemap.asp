<!--#include virtual="/shared/inc_access.asp" -->
<html><!-- #BeginTemplate "/Templates/2cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Ecoinfo</title>
<link href="/shared/styles.css" rel="stylesheet" type="text/css">
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
<!-- START HEADER --><!-- #BeginLibraryItem "/Library/header.lbi" -->
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
  <table width="90%"  border="0" align="center">
    <tr>
      <td><p>&nbsp;</p>
        <p class="contentHeader1">Sitemap</p>
        <p>som &Oslash;ko-net ser det </p>
        <p>Dette sitemap vises kun for &Oslash;ko-net medarbejdere</p>
        <p><a href="/shared/help/help.asp">Hj&aelig;lp til brug af funktionerne </a> </p></td>
    </tr>
  </table>
  <!-- Sitemap menu kan kopieres mellem alle de n&aelig;vnte sites -->
<!-- end sitemap menu -->
<!-- #EndEditable --></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="569" valign="top" height="100%" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> <!-- #BeginEditable "maincontent" -->
<!-- Sitemap kan kopieres mellem alle de n&aelig;vne sites -->
<table width="100%"  border="0">
  <tr>
    <td height="540">
      <table width="95%"  border="0" align="center">
        <tr valign="middle">
          <td width="30%"><span class="contentHeader1">&Oslash;ko-info Insider</span></td>
          <td width="30%"><span class="contentHeader1"> </span> <span class="sitetag"><a href="http://insider.eco-info.dk" class="contentHeader2">insider.eco-info.dk</a></span></td>
          <td width="30%" class="sitetag">Administration af data der vises p&aring; &Oslash;ko-info m.v. </td>
        </tr>
      </table>
      <ul>
        <li><span class="contentHeader2"><a href="http://www.eco-info.dk/index.asp">Forside</a></span> <span class="sitetag">med login</span> <span class="sitetag">og statistik
            </span>
          <ul>
              <li><span class="contentHeader2"><a href="/dgs/index.asp">Organisationer</a></span>
                  <ul>
                    <li><a href="/dgs/list.asp">S&oslash;gbar liste</a> <span class="sitetag">over organisation</span> </li>
                    <li>Rediger <span class="sitetag">organisation</span></li>
                    </ul>
              </li>
              <li><span class="contentHeader2"><a href="/ok/index.asp">Arrangementer</a></span>
                  <ul>
                    <li><a href="/ok/list.asp">S&oslash;gbar liste</a> <span class="sitetag">over  arrangementer</span></li>
                    <li><a href="/dgs/list.asp">Nyt</a> <span class="sitetag">arrangement oprettes ud fra organisation </span></li>
                    <li>Rediger <span class="sitetag">arrangement</span> </li>
                  </ul>
              </li>
              <li><span class="contentHeader2"><a href="/dgb/list.asp">Publikationer</a></span>
                  <ul>
                    <li><a href="/dgb/list.asp">S&oslash;gbar liste</a> <span class="sitetag">over  publikationer </span></li>
                    <li><a href="/dgs/list.asp">Ny</a> <span class="sitetag">publikation</span> <span class="sitetag"> oprettes ud fra organisation</span></li>
                    <li>Rediger <span class="sitetag">publikation</span></li>
                  </ul>
              </li>
              <li><span class="contentHeader2"><a href="/kurser/list.asp">Kurser</a></span>
                  <ul>
                    <li><a href="/kurser/list.asp">S&oslash;gbar liste</a> <span class="sitetag">over kurser</span> </li>
                    <li><a href="/dgs/list.asp">Nyt</a> <span class="sitetag">kursus</span> <span class="sitetag"> oprettes ud fra organisation</span></li>
                    <li>Rediger <span class="sitetag">kursus</span></li>
                    <li><span class="contentHeader2"><a href="/udd/list.asp">Uddannelser</a></span>
                        <ul>
                          <li><a href="/udd/list.asp">S&oslash;gbar liste</a> <span class="sitetag">over uddannelser</span></li>
                          <li><a href="/dgs/list.asp">Ny</a> <span class="sitetag">uddannelse oprettes ud fra organisation</span></li>
                          <li>Rediger <span class="sitetag">uddannelse</span> </li>
                        </ul>
                    </li>
                  </ul>
              </li>
              <li><span class="contentHeader2"><a href="/art/list.asp">Artikler</a></span>
                  <ul>
                    <li><a href="/art/list.asp">S&oslash;gbar liste</a> <span class="sitetag">over artikler</span></li>
                    <li><a href="/dgs/list.asp">Ny</a> <span class="sitetag">artikel oprettes ud fra organisation</span></li>
                    <li>Rediger <span class="sitetag">artikel</span></li>
                  </ul>
              </li>
              <li><span class="contentHeader2"><a href="http://insider.eco-info.dk/home/about_1.asp">Aktuelt</a></span>
                  <ul>
                    <li><a href="/news/list.asp">S&oslash;gbar liste</a> <span class="sitetag">over nyheder</span> </li>
                    <li><a href="/news/new.asp">Ny nyhed</a> </li>
                    <li>Rediger <span class="sitetag">nyhed</span> </li>
                    </ul>
              </li>
              <li><span class="contentHeader2"><a href="/admin/ei/vindue/functions.asp">ADM</a></span>
                <span class="sitetag">administration af hjemmesider                
                </span>
                <ul>
                    <li><span class="formLabeltext"><a href="/admin/ei/homepage/functions.asp">&Oslash;ko-info</a></span>                      
                      <ul>
                        <li><a href="/admin/ei/vindue/functions.asp">Vindue</a>  <span class="sitetag">til forside og vinduesside</span> 
                          <ul>
                            <li>rediger leder</li>
                            <li>rediger  artikler</li>
                            <li><a href="/admin/ei/vindue/functions.asp">rediger sektioner til forside </a>
                              <ul>
                                <li><a href="/admin/ei/vindue/edit_sec_dgs.asp">De Gr&oslash;nne Side</a></li>
                                <li><a href="/admin/ei/vindue/edit_sec_ok.asp">&Oslash;ko-kalender</a></li>
                                <li><a href="/admin/ei/vindue/edit_sec_dgb.asp">Det Gr&oslash;nne Bibliotek</a></li>
                                <li><a href="/admin/ei/vindue/edit_sec_opsl.asp">Opslagstavle</a></li>
                                <li><a href="/admin/ei/vindue/edit_sec_news.asp">Nyheder</a></li>
                              </ul>
                            </li>
                            </ul>
                        </li>
                        <li>Tema <span class="sitetag">bruges ikke</span> </li>
                        <li><a href="/admin/ei/homepage/functions.asp">Forsiden</a> <span class="sitetag">antal nyheder, splash og branding bruges ikke </span></li>
                        <li><a href="/admin/ei/sections/functions.asp">Sektioner</a> <span class="sitetag">startvisning bruges kun p&aring; De Gr&oslash;nne Sider</span> </li>
                        <li><a href="/admin/ei/cats/functions.asp">Kategorier og emneord</a> <span class="sitetag">redigeres til s&oslash;gning i sektionerne </span></li>
                        <li><a href="/admin/ei/ads/functions.asp">Annoncer og bannere</a> <span class="sitetag">oprettelse og valg for dgs, dgb og aktuelt</span></li>
                        <li><a href="/admin/ei/mails/functions.asp">Emails</a> <span class="sitetag">afsendelse til brugere</span> </li>
                        </ul>
                    </li>
                    <li><span class="formLabeltext">BU</span> <span class="sitetag">administrationen overg&aring;r til ny CMS</span> </li>
                    <li><span class="formLabeltext"><a href="/admin/en/index.asp">&Oslash;ko-net</a></span>
                      <ul>
                        <li><a href="/admin/en/index.asp">&Oslash;ko-net blad</a> <span class="sitetag">opret, v&aelig;lg blad</span> 
                          <ul>
                            <li>Rediger Forside</li>
                            <li>Rediger Afsnit</li>
                            <li>Rediger Artikel</li>
                            <li>Upload PDF    </li>
                          </ul>
                        </li>
                        <li><a href="/admin/en/forside.asp">Forside</a> <span class="sitetag">teaser, antal nyheder</span> </li>
                      </ul>
                    </li>
                    <li><span class="formLabeltext"><a href="/admin/mad/index.asp">Mad og musik</a></span>
                      <ul>
                        <li>Opret arrangement </li>
                        <li>Rediger arrangement</li>
                      </ul>
                    </li>
                    <li><span class="formLabeltext"><a href="/admin/ei/stat/functions.asp">Statistik</a></span> <span class="sitetag">bes&oslash;gende p&aring; hjemmesiderne</span> </li>
                    <li><span class="formLabeltext"><a href="/admin/ei/billedbasen/index.asp">Billedbasen</a></span>
                      <ul>
                        <li><a href="/admin/ei/billedbasen/index.asp">S&oslash;gbar liste</a> <span class="sitetag">over billeder p&aring; hjemmesiderne</span> </li>
                        <li><a href="/admin/ei/billedbasen/insert.asp">Opret nyt</a> <span class="sitetag">GIF billede</span></li>
                        <li><a href="/admin/ei/billedbasen/insert_jpg.asp">Opret nyt</a> <span class="sitetag">JPG billede</span></li>
                        <li><a href="/admin/ei/billedbasen/cat.asp">Kategorier</a> <span class="sitetag">redigeres til billederne</span> </li>
                      </ul>
                    </li>
                    <li><span class="formLabeltext"><a href="/admin/ei/request/index.asp">Foresp&oslash;rgsel</a></span>                      
                      <ul>
                        <li><a href="/admin/ei/request/orgmail.asp">Organisation uden email </a></li>
                        <li><a href="/admin/ei/request/orgdescr.asp">Organisation uden beskrivelse</a></li>
                        <li><a href="/admin/ei/request/orgshortdescr.asp">Organisation uden kort beskrivelse</a> </li>
                        <li><a href="/dgb/dblist.asp">Publikationer</a> <span class="sitetag">liste over alle</span>  </li>
                      </ul>
                    </li>
                    <li><span class="formLabeltext"><a href="/admin/ei/help/functions.asp">Hj&aelig;lpetekst</a></span> <span class="sitetag">redigering af Insiderhj&aelig;lp  hierarkisk</span>                 
                      <ul>
                        <li><a href="/admin/ei/help/functions.asp?intern=0">&Oslash;ko-net</a></li>
                        <li><a href="/admin/ei/help/functions.asp?intern=1">Bruger</a></li>
                      </ul>
                    </li>
                    </ul>
              </li>
            </ul>
        </li>
    </ul></td>
  </tr>
</table>
<!-- end sitemap -->
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
