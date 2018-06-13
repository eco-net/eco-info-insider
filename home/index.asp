<!--#include virtual="/shared/sqlcheck.asp"-->
<html>
<!-- #BeginTemplate "/Templates/3cols.dwt" -->
<head>
<!-- #BeginEditable "doctitle" -->
<title>Øko-info Insider</title>
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
    <td width="750" valign="top"><!-- START HEADER -->
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
          <td height="17" align="right" colspan="4"><table border="0" cellpadding="0" cellspacing="0" width="100%">
              <tr>
                <td align="left"><img src="/shared/graphics/logo2.gif" width="21" height="16"></td>
                <td align="right"><a href="/home/index.asp">Forside</a> | <a href="http://www.eco-info.dk" target="_blank">&Oslash;ko-info</a> | <a href="/home/about_1.asp">Om Øko-info Insider</a> | <a href="/home/searching.asp">Kom 
                  i gang</a> </td>
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
          <td width="388" height="64"><div align="center"> </div>
          </td>
          <td background="/shared/graphics/layout/dots_vert.gif" width="1"><br>
          </td>
          <td width="180" align="center" valign="top" background="/shared/graphics/layout/search_bkgrd.gif"><table width="150" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
              <tr>
                <td></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <!-- #EndLibraryItem -->
      <!-- END HEADER -->
      <!-- #BeginEditable "menu" -->
      <!--include virtual="/shared/stat.asp" -->
      <%
DIM curtab
curtab=0%>
      <!-- #BeginLibraryItem "/Library/menu.lbi" -->
      <table width="750" border="0" cellspacing="0" cellpadding="0" name="Menu">
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
	'-- tab 5 --
	
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
	END IF
	response.write "<img src=""/shared/graphics/menu/stretch.gif"" width=""164"" height=""22""></td>"
	
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
        <tr>
          <td class="menuBar"><%
SET fs = CreateObject("Scripting.FileSystemObject")
Set ts=fs.OpenTextFile(Server.mapPath("inc_submenu.txt"))
execute(ts.ReadAll())
ts=""
fs=""
%>
          </td>
        </tr>
        <%END IF%>
        <tr>
          <td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
        </tr>
      </table>
      <!-- #EndLibraryItem --><!-- #EndEditable -->
      <table width="750" border="0" cellspacing="0" cellpadding="0" name="Contentarea">
        <tr>
          <td width="180" valign="top" name="leftbar"><!-- #BeginEditable "leftbar" -->
            <table width="90%"  border="0" align="center">
              <tr>
                <td height="30" class="sidebarHeader"><div align="center"><a href="sitemap.asp"> Sitemap<img src="/shared/graphics/layout/folder.gif" alt="Alle &Oslash;ko-net's sites" width="15" height="15" hspace="5" border="0"></a><br>
                  </div></td>
              </tr>
            </table>
            <%IF LEN(request.cookies("eiuserid"))>0 THEN%>
            <table width="90%"  border="0" align="center">
              <tr>
                <td height="30" class="sidebarHeader"><div align="center">
                    <p><a href="insider_sitemap.asp"> Sitemap<img src="/shared/graphics/layout/folder.gif" alt="Alle &Oslash;ko-net's sites" width="15" height="15" hspace="5" border="0"></a><br>
                      for &Oslash;ko-net </p>
<!-- <p><a href="handbook/index.asp">Webmasterens h&aring;ndbog <img src="/shared/graphics/layout/book.jpg" width="26" height="24" border="0"></a><br> </p>-->
                  </div></td>
              </tr>
            </table>
            <!--include file="membercount.asp" -->
            <table width="180" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1" /></td>
              </tr>
              <tr>
                <td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22" /></td>
                <td width="98%" class="sidebarHeader">&nbsp;&nbsp;NYE MEDLEMMER </td>
              </tr>
              <tr>
                <td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1" /></td>
              </tr>
              <tr>
                <td valign="top" colspan="2"><table width="90%" border="0" align="center" cellpadding="5">
                    <tr>
                      <td><form name="form3" method="post" action="new_members.asp">
                          <div align="center">Nye betalende medlemmer idag: <br>
                            <input name="members" type="text" id="members" value="0" size="2">
                            &nbsp;&nbsp;
                            <input name="Registrer" type="submit" class="formbutton" id="Registrer" value="Registrer">
                            <br>
                            Der er reg. ialt: <%=mAntal%><br>
                            sidste gang <%=mDato%></div>
                        </form></td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <% end if %>
            <!--#include virtual="/shared/inc_login.asp"-->
            <!-- #EndEditable --></td>
          <td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
          </td>
          <td width="388" height="100%" valign="top" name="maincontent"><table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td valign="top"><!-- #BeginEditable "maincontent" -->
                  <div align="center">
                    <%IF LEN(request.cookies("eiuserid"))=0 THEN%>
                    <!--#include file="welcome.asp"-->
                    <p>
                      <%else %>
					  <span class="contentHeader1">Google Analytics</span>:</p>
                    <p>log p&aring; Google med web@eco-net.dk / ecoimages </p>
                    <p><a href="stat.asp">Hitlisten</a>
                      <!--#include file="ecostat.asp"-->
                      <%end if %>
                        </p>
                  </div>
                  <!-- #EndEditable --> </td>
              </tr>
              <tr>
                <td valign="bottom" align="left"><!--#include virtual="/shared/pagetools.asp"-->
                </td>
              </tr>
            </table></td>
          <td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
          </td>
          <td width="180" valign="top" name="rightbar"><!-- #BeginEditable "rightbar" --><!-- #BeginLibraryItem "/Library/Sektioner.lbi" -->
            <%
IF LEN(request.cookies("eiorgid"))>0 THEN
	thepage="index.asp"

ELSEIF LEN(request.cookies("eiuserid"))>0 THEN
	thepage="list.asp?valid=0"
ELSE
	thepage="index.asp"
END IF
%>
            <table width="180" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
                <td width="98%" class="sidebarHeader">&nbsp;&nbsp;SEKTIONER</td>
              </tr>
              <tr>
                <td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
              </tr>
              <tr valign="top">
                <td colspan="2" class="sidebarText"><table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
                    <tr>
                      <td><img src="/shared/graphics/spacer.gif" width="3" height="3"></td>
                    </tr>
                    <tr>
                      <td><span class="sidebarHeader">Stamdata</span><br>
                        Jeres/dine oplysninger i De gr&oslash;nne sider.<br>
                        Det er ogs&aring; her du indstiller brugen af data p&aring; Jeres/din egen site.<br>
                        <br>
                      </td>
                    </tr>
                    <tr>
                      <td><span class="sidebarHeader">Arrangementer</span><br>
                        Her opretter og &aelig;ndre du Jeres/dine arrangementer i &Oslash;ko-kalenderen.<br>
                        <br>
                      </td>
                    </tr>
                    <tr>
                      <td><span class="sidebarHeader">Publikationer</span><br>
                        Her opretter og &aelig;ndrer du jeres/dine publikationer, links til viden p&aring; 
                        Jeres/din site mv.<br>
                        <br>
                      </td>
                    </tr>
                    <tr>
                      <td><span class="sidebarHeader">Kurser</span><br>
                        Her er en kursusdatabase under udvikling<br>
                        <br>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            <!-- #EndLibraryItem -->
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
              </tr>
              <tr>
                <td><img src="/shared/graphics/layout/header_arrows.gif" width="22" height="22"></td>
                <td width="98%" class="sidebarHeader">&nbsp;&nbsp;ANDRE &Oslash;KO-NET SIDER</td>
              </tr>
              <tr>
                <td colspan="2" height="1" background="/shared/graphics/layout/dots_horz.gif"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
              </tr>
              <tr>
                <td colspan="2"><div align="center"><a href="http://www.eco-info.dk"><br>
                    </a>
                    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
                      <tr>
                        <td><p>&nbsp;</p>
                          <p align="center" class="listheader"><a href="http://www.sustainabledevelopment.dk">Sustainable Development</a></p>
                          <p align="center" class="listheader"><a href="http://www.eco-info.dk">&Oslash;ko-info</a></p>
                          <p align="center" class="listheader"><a href="http://www.eco-net.dk">&Oslash;ko-net</a></p>
                          <p align="center" class="listheader">&nbsp;</p></td>
                      </tr>
                    </table>
                  </div>
                  <div align="center"><br>
                  </div>
                  <div align="center"><a href="http://www.bæredygtigudvikling.nu"><br>
                    </a> <a href="#" onClick="window.open('/shared/help/help.asp','Help','scrollbars=yes,resizable=yes,width=700,height=300');">Hj&aelig;lp</a><br>
                  </div></td>
              </tr>
            </table>
            <!-- #EndEditable --></td>
        </tr>
      </table></td>
    <td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
  </tr>
  <tr>
    <td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="3"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
  </tr>
  <tr align="center">
    <td colspan="3" class="footer" height="20" valign="middle"><!--#include virtual="/shared/footer.asp"-->
    </td>
  </tr>
</table>
</body>
<!-- #EndTemplate -->
<script src="http://www.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "UA-3636900-6";
urchinTracker();
</script>
</html>
<!--include virtual="/shared/log.asp"-->
<%'reg("homeindex")%>
