<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Returmail</title>
<style type="text/css">
<!--
.style1 {font-family: Verdana, Arial, Helvetica, sans-serif}
-->
</style>
</head>

<body>
<table width="90%" border="0" align="center" cellpadding="5" cellspacing="5" class="style1">
  <tr>
    <td colspan="2"><div align="center" class="style1">
      <h2>Returmail </h2>
    </div></td>
  </tr>
  <tr>
    <td><p><strong>Procedure </strong></p>    </td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td valign="top"><p>Udv&aelig;lgelse og afsendelse</p>    </td>
    <td valign="top"><ol>
      <li>Lars udv&aelig;lger mailadresser fra Filemaker til nyhedsmail</li>
      <li>Lars uploader disse mailadresser til nettet: tabellen eisys_fmpmail </li>
    </ol></td>
  </tr>
  <tr>
    <td valign="top">K&oslash;rsel inden ny afsendelse</td>
    <td valign="top"><p>Tabellen eisys_returnmail indeholder mailadresser der er returneret ved sidste afsendelse af nyhedsmail.</p>
    <p>Ved at k&oslash;re et program der sletter mailadresser fra  eisys_fmpmail som ogs&aring; er i eisys_returnmail, fjernes de overfl&oslash;dige inden afsendelse af nyhedsmail.</p>
    <p><a href="/admin/ei/mails/returmail/remove.asp">K&oslash;r fjernelse af overfl&oslash;dige</a>  </p></td>
  </tr>
  <tr>
    <td valign="top">Behandling af returmails</td>
    <td valign="top"><p>Efter afsendelse af nyhedsmail hentes returmails ind i Outlook: nyhedsmail@eco-net.dk</p>
    <p>Disse mails eksporteres i Outlook til Access:</p>
    <ol>
      <li>Filer / Importer og eksporter </li>
      <li>Eksporter til en fil</li>
      <li>Access</li>
      <li>Indbakke</li>
      <li>Indbakke.mdb </li>
      </ol>
    <p>Disse mails skal nu over i en tekstfil: </p>
    <ol>
      <li>I Access, Indbakke.mdb tabellen Email kopieres kolonnen Meddelelsestekst.</li>
      <li>S&aelig;t Ind i Notesblok</li>
      <li>Kopier alt i Notesblok og S&aelig;t ind i Word (hvis der er mange) </li>
      <li>Erstat &quot; med # i Word: Rediger/Erstat (hvis der er f&aring; kan Notesblok ogs&aring;) </li>
      <li>Kopier alt i Word og S&aelig;t ind i filen returmails.txt i Insider:<br />
        admin/ei/mails/returmail/returmails.txt </li>
      <li>Upload returmails.txt til serveren med Dreamweaver </li>
    </ol></td>
  </tr>
  <tr>
    <td valign="top">Retur mailadresser til database </td>
    <td valign="top"><p>Ved at k&oslash;re et program der l&aelig;ser de emailadreseer der befinder sig i meddelelsesteksterne, og sammenligne med de afsendte fra tabellen eisys_fmpmail, overskrives tabellen eisys_returnmail med de relevante returmailadresser.</p>
      <p><a href="returmail.asp">K&oslash;r overskrivelse af tabellen eisys_returnmail</a></p></td>
  </tr>
  <tr>
    <td valign="top">&nbsp;</td>
    <td valign="top">&nbsp;</td>
  </tr>
</table>
</body>

</html>
