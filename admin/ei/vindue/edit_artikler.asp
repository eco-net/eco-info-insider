<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/shared/common.asp" -->

<%
page_id=request("id")
%>

<%
set rsVin = Server.CreateObject("ADODB.Recordset")
rsVin.ActiveConnection = MM_ecoinfo_STRING
rsVin.CursorType = 0
rsVin.CursorLocation = 2
rsVin.LockType = 3
rsVin.Source = "SELECT * FROM eivindue_item WHERE page_id=" & page_id & " ORDER BY thesort"
rsVin.Open()
%>
<%
set rsName = Server.CreateObject("ADODB.Recordset")
rsName.ActiveConnection = MM_ecoinfo_STRING
rsName.CursorType = 0
rsName.CursorLocation = 2
rsName.LockType = 3
rsName.Source = "SELECT * FROM eivindue_maindata WHERE id=" & page_id
rsName.Open()
name=rsName("name")
rsName.Close()
%>

<script language="javascript">
function del_item(itemid,pageid){
 if (confirm("Vil du nu virkelig slette, eller løb musen løbsk?")) {
 location.href="del_item.asp?item_id=" + itemid + "&page_id=" + pageid;
}
}
</script>
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
<!-- #BeginEditable "menu" --> 
<%DIM curtab
curtab=6
level1=3
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
<!--#include virtual="/shared/showimage.asp" -->
  <table width="180" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="180" height="35" bgcolor="#D1CEC9"><a href="edit.asp?id=<%=page_id%>&name=<%=name%>">&lt;&lt; Tilbage til hovedmenu </a></td>
    </tr>
	<% if not rsVin.EOF then
	  while not rsVin.EOF
	  %>
    <tr>
      <td valign="top">
	  <% if rsVin("col")=1 then %>
	  <% if rsVin("img_id")=0 then %>
	  	 <%=rsVin("tekst")%><br><br>
		  <% if rsVin("source")<>"" then %>
	<i>Kilde: <%=rsVin("source")%></i><br><br>
	<% end if %>
	<a href="<%=rsVin("link")%>"><%=rsVin("linktext")%></a><br>
		<b><%=rsVin("thesort")%></b>
		<a href="/admin/edit_vindue_text.asp?item_id=<%=rsVin("id")%>&name=<%=request("name")%>&page_id=<%=page_id%>"><img src="/shared/graphics/layout/edit.gif" alt="Rediger" width="15" height="13" border="0"> 
        </a><a href="javascript:del_item(<%=rsVin("id")%>,<%=page_id%>)"><img src="/shared/graphics/layout/delete.jpg" alt="Slet" width="16" height="15" border="0"></a>        
		   <% else %>	
			<div align="center">
		    <img src="<%=getImage(rsVin("img_id"),"R")%>"> <br>
		    <% =rsVin("img_text") %>
		    <br> </div>
			<b><%=rsVin("thesort")%></b>
		    <a href="edit_img.asp?item_id=<%=rsVin("id")%>&name=<%=request("name")%>&page_id=<%=page_id%>"><img src="/shared/graphics/layout/edit.gif" alt="Rediger" width="16" height="13" border="0"> </a><a href="javascript:del_item(<%=rsVin("id")%>,<%=page_id%>);"><img src="/shared/graphics/layout/delete.jpg" alt="Slet" width="16" height="15" border="0"></a>
            <% end if %>
			<hr>
	        <% end if %>	     </td>
    </tr>
<% rsVin.MoveNext
	  wend
	    end if 
	   rsVin.Close()%>
  </table>
  <br>
<!-- #EndEditable --></td>
<td width="1" background="/shared/graphics/layout/dots_vert.gif"><br>
</td>
<td width="569" valign="top" height="100%" name="maincontent"> 
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td valign="top"> <!-- #BeginEditable "maincontent" --> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="382" height="35" bgcolor="#D1CEC9"><div align="center"><span class="contentHeader2"><strong>Rediger Artikel til Vindue:</strong> <%=name%></span> </div></td>
    <td width="1" background="/shared/graphics/layout/dots_vert.gif"></td>
    <td width="180" height="35" bgcolor="#D1CEC9"><strong>Inds&aelig;t nyt:</strong> <a href="/admin/new_vindue_text.asp?page_id=<%=request("id")%>&name=<%=request("name")%>"><img src="/shared/graphics/layout/text.gif" alt="Inds&aelig;t ny tekst" width="29" height="24" hspace="5" border="0"></a><a href="new_img.asp?page_id=<%=request("id")%>&name=<%=request("name")%>"><img src="/shared/graphics/layout/img.gif" alt="Inds&aelig;t billede fra billedbasen" width="29" height="24" hspace="5" border="0"></a></td>
  </tr>
  <tr>
    <td height="35">
      <table width="100%" border="0" align="center">
        <tr>
          <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
              <% rsVin.Open() 
		if not rsVin.EOF then
	  while not rsVin.EOF
	  %>
              <tr>
                <td height="100%" valign="top">
                  <% if rsVin("col")=2 then %>
                  <% if rsVin("img_id")=0 then %>
                  <%=rsVin("tekst")%><br><br>
				  <% if rsVin("source")<>"" then %>
	<i>Kilde: <%=rsVin("source")%></i><br><br>
	<% end if %>
	<a href="<%=rsVin("link")%>"><%=rsVin("linktext")%></a><br>
				  <b><%=rsVin("thesort")%></b>
                  <a href="/admin/edit_vindue_text.asp?item_id=<%=rsVin("id")%>&name=<%=request("name")%>&page_id=<%=page_id%>"><img src="/shared/graphics/layout/edit.gif" alt="Rediger" width="15" height="13" border="0"> </a><a href="javascript:del_item(<%=rsVin("id")%>,<%=page_id%>)"><img src="/shared/graphics/layout/delete.jpg" alt="Slet" width="16" height="15" border="0"></a>
                    <% else %>
					  <div align="center">
                      <img src="<%=getImage(rsVin("img_id"),"3")%>"><br>
  				    <br>
                      <% =rsVin("img_text") %></div>
                      <br><b><%=rsVin("thesort")%></b>
                      <a href="edit_img.asp?item_id=<%=rsVin("id")%>&name=<%=request("name")%>&page_id=<%=page_id%>"><img src="/shared/graphics/layout/edit.gif" alt="Rediger" width="16" height="13" border="0"> </a><a href="javascript:del_item(<%=rsVin("id")%>,<%=page_id%>);"><img src="/shared/graphics/layout/delete.jpg" alt="Slet" width="16" height="15" border="0"></a>
                      <% end if %>
					  <hr>
                      <% end if %>
                  </div></td>
              </tr>
              <% rsVin.MoveNext
	  wend
	 
	   end if 
	    rsVin.Close()%>
          </table></td>
        </tr>
    </table></td>
    <td background="/shared/graphics/layout/dots_vert.gif"></td>
    <td valign="top"><table width="100%"  border="0" align="center">
        <tr>
          <td><table width="180" border="0" align="center" cellpadding="0" cellspacing="0">
              <% rsVin.Open()
		if not rsVin.EOF then
	  while not rsVin.EOF
	  %>
              <tr>
                <td valign="top">
                  <% if rsVin("col")=3 then %>
                  <% if rsVin("img_id")=0 then %>
                   <%=rsVin("tekst")%><br><br>
				    <% if rsVin("source")<>"" then %>
	<i>Kilde: <%=rsVin("source")%></i><br><br>
	<% end if %>	<a href="<%=rsVin("link")%>"><%=rsVin("linktext")%></a><br>
				  <b><%=rsVin("thesort")%></b>
                  <a href="/admin/edit_vindue_text.asp?item_id=<%=rsVin("id")%>&name=<%=request("name")%>&page_id=<%=page_id%>"><img src="/shared/graphics/layout/edit.gif" alt="Rediger" width="15" height="13" border="0"> </a><a href="javascript:del_item(<%=rsVin("id")%>,<%=page_id%>)"><img src="/shared/graphics/layout/delete.jpg" alt="Slet" width="16" height="15" border="0"></a>
                      <% else %>
					   <div align="center"> <img src="<%=getImage(rsVin("img_id"),"R")%>"> <br>
                      <% =rsVin("img_text") %>
                      <br> </div>
					  <b><%=rsVin("thesort")%></b>
                       <a href="edit_img.asp?item_id=<%=rsVin("id")%>&name=<%=request("name")%>&page_id=<%=page_id%>"><img src="/shared/graphics/layout/edit.gif" alt="Rediger" width="16" height="13" border="0"> </a><a href="javascript:del_item(<%=rsVin("id")%>,<%=page_id%>);"><img src="/shared/graphics/layout/delete.jpg" alt="Slet" width="16" height="15" border="0"></a>
                      <% end if %>
					  <hr>
                      <% end if %>                 </td>
              </tr>
              <% rsVin.MoveNext
	  wend
	   end if 
	   rsVin.Close()%>
          </table></td>
        </tr>
    </table></td>
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
<!--#include virtual="/shared/log.asp"-->
