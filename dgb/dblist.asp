
<!--#include virtual="/Connections/ecoinfo.asp" -->

<%

set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT m.id,m.title,m.author,lo.catid,m.isbn,m.language,m.editor,m.islocal,m.year FROM eilib_maindata m  ORDER BY m.title"
'rsPageData.Source=replace(rsPageData.Source,"0=0",theWhere)
theFrom= "(eilib_maindata m LEFT JOIN eilib_r_cat lo ON m.id=lo.libid)"
rsPageData.Source=replace(rsPageData.Source,"eilib_maindata m",theFrom)
if request("sort")<>"" then
thesort="ORDER BY " & request("sort")
rsPageData.Source=replace(rsPageData.Source,"ORDER BY m.title",thesort)
end if
'response.write rsPageData.Source
'response.end
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0

Function catname(cno)
if cno=182 then catname="Årbog" end if
if cno=165 then catname="Artikel" end if
if cno=11 then catname="Bog" end if
if cno=297 then catname="CD-medier" end if
if cno=333 then catname="hjemmeside" end if
if cno=234 then catname="Film,video" end if
if cno=12 then catname="Årbog" end if
if cno=294 then catname="Hæfte,pjece" end if
if cno=359 then catname="Internet" end if
if cno=360 then catname="Miljø-lokal" end if
if cno=182 then catname="Miljø-national" end if
if cno=165 then catname="Radio/TV" end if
if cno=11 then catname="Rapport" end if
if cno=297 then catname="Spil" end if
if cno=333 then catname="Tema-nummer" end if
if cno=234 then catname="Tidsskrift" end if
end function
%>
<html>
<head>
<title>Alle Publikationer</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.style3 {font-size: 11px; font-family: Verdana, Arial, Helvetica, sans-serif; }
.style5 {font-size: 11px; font-family: Verdana, Arial, Helvetica, sans-serif; font-weight: bold; }
-->
</style>
</head>

<body>
<table width="100%"  border="1">
 <tr>
    <td class="style3"><strong><a href="dblist.asp?sort=title">Titel</span></a></strong></td>
	<td class="style3"><strong><a href="dblist.asp?sort=author">Forfatter</a></span></strong></td>
    <td><span class="style5"><a href="dblist.asp?sort=catid">Kategori</a></span></td>
    <td><span class="style5"><a href="dblist.asp?sort=isbn">ISBN</a></span></td>
    <td><span class="style5"><a href="dblist.asp?sort=language">Sprog</a></span></td>
    <td><span class="style5"><a href="dblist.asp?sort=editor">Udgiver</a></span></td>
    <td><span class="style5"><a href="dblist.asp?sort=islocal">Økonet</a></span></td>
	<td><span class="style5"><a href="dblist.asp?sort=year">År</a></span></td>
  </tr>
<% while not rsPageData.EOF %>
  <tr>
    <td><span class="style3"><a href="javascript:opener.location.href='edit.asp?id=<%=rsPageData("id")%>';this.history.go(-1)"><%=Left(rsPageData("title"),25)%></a></span></td>
	 <td><span class="style3"><%=Left(rsPageData("author"),25)%>&nbsp;</span></td>
    <td><span class="style3"><%=catname(rsPageData("catid"))%>&nbsp;</span></td>
    <td><span class="style3"><%=rsPageData("isbn")%>&nbsp;</span></td>
    <td><span class="style3"><%=rsPageData("language")%>&nbsp;</span></td>
    <td><span class="style3"><%=Left(rsPageData("editor"),15)%>&nbsp;</span></td>
    <td><span class="style3"><% if rsPageData("islocal")>0 then response.write "Ø" %>&nbsp;</span></td>
	<td><span class="style3"><%=rsPageData("year")%>&nbsp;</span></td>
  </tr>
  <%
rsPageData.MoveNext
wend
%>
</table>

<br>

<%
rsPageData.Close
%>

</body>
</html>
