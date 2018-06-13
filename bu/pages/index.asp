<!--#include virtual="/bu/snippets/news.asp"-->
<!--#include virtual="/bu/snippets/cal.asp"-->
<!--#include virtual="/bu/snippets/debate.asp"-->
<!--#include virtual="/bu/snippets/blog.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>B&aelig;redygtig Udvikling</title>
<link href="/bu/common/style.css" rel="stylesheet" type="text/css" media="screen" />
</head>

<body>
<table width="962" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="64" colspan="10" background="/bu/images/header-bg.gif"><a href="index.asp"><img src="/bu/images/logo.gif" alt="gr&oslash;nne sider" width="454" height="64" border="0" /></a></td>
  </tr>
  <tr>
  <td><table width="100%" height="36" border="0" cellpadding="0" class="headermenu">
    <tr>
      <td width="40">&nbsp;</td>
      <td width="80"><div align="center"><a href="index.asp">Forside</a></div></td>
      <td width="80"><div align="center"><span class="style5">Om BU </span></div></td>
      <td width="80"><div align="center"><span class="style5">Login</span></div></td>
      <td width="80"><div align="center"><span class="style5">Blog</span></div></td>
      <td width="80"><div align="center"><span class="style5">St&oslash;t</span></div></td>
      <td width="80"><div align="center"><span class="style5">FAQ</span></div></td>
      <td>&nbsp;</td>
      </tr>
  </table></td>
  </tr>
  <tr>
    <td height="500" colspan="10" valign="top" bgcolor="#ffffff"><br />
      <table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="255" valign="top"><p class="big">Alt om BU </p>
          <ul class="navmenu">
            <li><a href="/designtest/bu/hvad_er_BU.html">Introduktion</a></li>
            <li>BU-internationalt</li>
            <li>Bu i Danmark</li>
            <li>Agenda 21</li>
            <li>Indikatorer</li>
            <li>Temaer</li>
            <li>Ressourcer  </li>
            </ul>
          <p class="big">Kampagner</p>
          <ul class="navmenu">
            <li>Balanceakten</li>
            <li>UBU-portalen</li>
            <li>CD-B&aelig;reklang</li>
            </ul>
          <p class="big">Vores Debat</p>
          <ul class="navmenu">
            <li>Blogs</li>
            <li>Fora </li>
          </ul>          
          <p class="big">S&oslash;g</p>
          <form id="form1" name="form1" method="post" action="">
            <label>
            <input name="textfield" type="text" class="searchbox" />
            </label>
            </form>          <p>&nbsp;</p></td>
        <td width="255" valign="top"><p class="big">Aktuelt</p>
          <p><%=getNews()%></p>
          <p class="big">Vores Debat </p>
          
          <p><%=getDebate()%></p></td>
        <td width="255" valign="top"><p class="big">Det Sker</p>
          <p><%=getCal()%>
          </p>
          <p class="big">Blogs</p>          
          <p><%=getBlog()%></p>
          <p>&nbsp;</p></td>
        <td width="180" valign="top">
		<script type="text/javascript"><!--
google_ad_client = "pub-1814652264778642";
/* 160x600, oprettet 02-04-08 */
google_ad_slot = "9452473333";
google_ad_width = 160;
google_ad_height = 600;
//-->
</script>
<script type="text/javascript"
src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="10" bgcolor="#ffffff"><div align="center">
      <p><span class="footer">b&aelig;redygtig udvikling   er etableret af <a href="http://www.eco-net.dk">&Oslash;ko-net</a>. &copy;   1997-2008 &Oslash;ko-net. Alle rettigheder forbeholdes.<br />
  Bes&oslash;g ogs&aring; portalen gr&oslash;&oslash;ne sider:<a href="http://www.grønnesider.dk" target="_blank"> www.gr&oslash;nnesider.dk</a><br />
        Kontakt: 62 24 43 24 | <a href="mailto:eco-net@eco-net.dk">eco-net@eco-net.dk</a></span></p>
    </div></td>
  </tr>
</table>
</body>
</html>
