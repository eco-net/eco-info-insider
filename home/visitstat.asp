
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
set rsStat = Server.CreateObject("ADODB.Recordset")
rsStat.ActiveConnection = MM_ecoinfo_STRING
rsStat.Source = "SELECT Count(*) as theCount  FROM eisys_insideruser_stat WHERE sek LIKE 'dgs' AND page=1"
rsStat.CursorType = 0
rsStat.CursorLocation = 2
rsStat.LockType = 3
rsStat.Open()
rsStat_numRows = 0
dgsnew=rsStat("theCount")
rsStat.Close
rsStat.Source = "SELECT Count(*) as theCount  FROM eisys_insideruser_stat WHERE sek LIKE 'dgs' AND page=2" 
rsStat.Open()
dgsedit=rsStat("theCount")
rsStat.Close
rsStat.Source = "SELECT Count(*) as theCount  FROM eisys_insideruser_stat WHERE sek LIKE 'dgb' AND page=1" 
rsStat.Open()
dgbnew=rsStat("theCount")
rsStat.Close
rsStat.Source = "SELECT Count(*) as theCount  FROM eisys_insideruser_stat WHERE sek LIKE 'dgb' AND page=2" 
rsStat.Open()
dgbedit=rsStat("theCount")
rsStat.Close
rsStat.Source = "SELECT Count(*) as theCount  FROM eisys_insideruser_stat WHERE sek LIKE 'ok' AND page=1" 
rsStat.Open()
oknew=rsStat("theCount")
rsStat.Close
rsStat.Source = "SELECT Count(*) as theCount  FROM eisys_insideruser_stat WHERE sek LIKE 'ok' AND page=2" 
rsStat.Open()
okedit=rsStat("theCount")
rsStat.Close
rsStat.Source = "SELECT Count(*) as theCount  FROM eisys_insideruser_stat WHERE sek LIKE 'kurser' AND page=1" 
rsStat.Open()
kurnew=rsStat("theCount")
rsStat.Close
rsStat.Source = "SELECT Count(*) as theCount  FROM eisys_insideruser_stat WHERE sek LIKE 'kurser' AND page=2" 
rsStat.Open()
kuredit=rsStat("theCount")
rsStat.Close
rsStat.Source = "SELECT Count(*) as theCount  FROM eisys_insideruser_stat WHERE sek LIKE 'udd' AND page=1" 
rsStat.Open()
uddnew=rsStat("theCount")
rsStat.Close
rsStat.Source = "SELECT Count(*) as theCount  FROM eisys_insideruser_stat WHERE sek LIKE 'udd' AND page=2" 
rsStat.Open()
uddedit=rsStat("theCount")
rsStat.Close
rsStat.Source = "SELECT Count(*) as theCount  FROM eisys_insideruser_stat WHERE sek LIKE 'art' AND page=1" 
rsStat.Open()
artnew=rsStat("theCount")
rsStat.Close
rsStat.Source = "SELECT Count(*) as theCount  FROM eisys_insideruser_stat WHERE sek LIKE 'art' AND page=1" 
rsStat.Open()
artedit=rsStat("theCount")
rsStat.Close
%>
<%
set rsCount = Server.CreateObject("ADODB.Recordset")
rsCount.ActiveConnection = MM_ecoinfo_STRING
rsCount.Source = "SELECT *  FROM eisys_insideruser_count"
rsCount.CursorType = 0
rsCount.CursorLocation = 2
rsCount.LockType = 3
rsCount.Open()
rsCount_numRows = 0
%>
<%


%>
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
<table width="380" border="0" cellspacing="0" cellpadding="0" background="/dgs/graphics/sub_index_header_dgs_bkgrd.gif" name="subIndexHeader" align="center">
<tr> 
<td valign="top"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td class="listheader" align="center"> <br>
<br>
Bes&oslash;g af brugere siden 25 nov. 2002</td>
</tr>
<tr> 
<td valign="top" class="sitetag"> 
<table width="100%" border="0" cellspacing="0" cellpadding="5">
<tr> 
<td width="15%" colspan="2" class="listheader"> 
<div align="center">dgs</div>
</td>
<td width="15%" colspan="2" class="listheader"> 
<div align="center">dgb</div>
</td>
<td width="15%" colspan="2" class="listheader"> 
<div align="center">ok</div>
</td>
<td width="15%" colspan="2" class="listheader"> 
<div align="center">kur</div>
</td>
<td width="15%" colspan="2" class="listheader"> 
<div align="center">udd</div>
</td>
<td width="15%" colspan="2" class="listheader"> 
<div align="center">art</div>
</td>
</tr>
<tr> 
<td width="7%"> 
<div align="center"><font color="#FF0000">new</font></div>
</td>
<td width="7%"> 
<div align="center"><font color="#0000FF">edit</font></div>
</td>
<td width="7%"> 
<div align="center"><font color="#FF0000">new</font></div>
</td>
<td width="7%"> 
<div align="center"><font color="#0000FF">edit</font></div>
</td>
<td width="7%"> 
<div align="center"><font color="#FF0000">new</font></div>
</td>
<td width="7%"> 
<div align="center"><font color="#0000FF">edit</font></div>
</td>
<td width="7%"> 
<div align="center"><font color="#FF0000">new</font></div>
</td>
<td width="7%"> 
<div align="center"><font color="#0000FF">edit</font></div>
</td>
<td width="7%"> 
<div align="center"><font color="#FF0000">new</font></div>
</td>
<td width="7%"> 
<div align="center"><font color="#0000FF">edit</font></div>
</td>
<td width="7%"> 
<div align="center"><font color="#FF0000">new</font></div>
</td>
<td width="7%"> 
<div align="center"><font color="#0000FF">edit</font></div>
</td>
</tr>
<tr> 
<td width="7%"> 
<div align="center"><%=dgsnew%></div>
</td>
<td width="7%"> 
<div align="center"><%=dgsedit%></div>
</td>
<td width="7%"> 
<div align="center"><%=dgbnew%></div>
</td>
<td width="7%"> 
<div align="center"><%=dgbedit%></div>
</td>
<td width="7%"> 
<div align="center"><%=oknew%></div>
</td>
<td width="7%"> 
<div align="center"><%=okedit%></div>
</td>
<td width="7%"> 
<div align="center"><%=kurnew%></div>
</td>
<td width="7%"> 
<div align="center"><%=kuredit%></div>
</td>
<td width="7%"> 
<div align="center"><%=uddnew%></div>
</td>
<td width="7%"> 
<div align="center"><%=uddedit%></div>
</td>
<td width="7%"> 
<div align="center"><%=artnew%></div>
</td>
<td width="7%"> 
<div align="center"><%=artedit%></div>
</td>
</tr>
</table>
</td>
</tr>
<tr> 
<td class="contentHeader1"> </td>
</tr>
</table>
</td>
</tr>
</table>
<br>
<table width="380" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
<br>
<br>
<table width="380" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td class="listheader" width="40%">Hits p&aring; udvalgte sider</td>
<td class="listheader" colspan="2"> 
<div align="center">idag</div>
</td>
<td class="listheader" colspan="2"> 
<div align="center">fra 25/11-02</div>
</td>
</tr>
<tr> 
<td width="40%">&nbsp;</td>
<td> 
<div align="center"><font color="#FF0000">bruger</font></div>
</td>
<td> 
<div align="center"><font color="#0000FF">&oslash;ko-net</font></div>
</td>
<td> 
<div align="center"><font color="#FF0000">bruger</font></div>
</td>
<td> 
<div align="center"><font color="#0000FF">&oslash;ko-net</font></div>
</td>
</tr>
<tr> 
<td width="40%">home/index</td>
<td> 
<div align="center"><%=Application("homeindex")%></div>
</td>
<td> 
<div align="center"><%=Application("ehomeindex")%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("homeindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("ehomeindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">dgs/list</td>
<td> 
<div align="center"><%=Application("dgslist")%></div>
</td>
<td> 
<div align="center"><%=Application("edgslist")%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("dgslist").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("edgslist").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">dgb/list</td>
<td> 
<div align="center"><%=Application("dgblist")%></div>
</td>
<td> 
<div align="center"><%=Application("edgblist")%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("dgblist").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("edgblist").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">ok/list</td>
<td> 
<div align="center"><%=Application("oklist")%></div>
</td>
<td> 
<div align="center"><%=Application("eoklist")%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("oklist").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("eoklist").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">kurser/list</td>
<td> 
<div align="center"><%=Application("kurlist")%></div>
</td>
<td> 
<div align="center"><%=Application("ekurlist")%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("kurlist").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("ekurlist").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">udd/list</td>
<td> 
<div align="center"><%=Application("uddlist")%></div>
</td>
<td> 
<div align="center"><%=Application("euddlist")%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("uddlist").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("euddlist").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">art/list</td>
<td> 
<div align="center"><%=Application("artlist")%></div>
</td>
<td> 
<div align="center"><%=Application("eartlist")%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("artlist").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("eartlist").Value)%></div>
</td>
</tr>
<tr> 
<td class="listheader" width="40%">Sessions</td>
<td> 
<div align="center"></div>
</td>
<td> 
<div align="center"></div>
</td>
<td> 
<div align="center"></div>
</td>
<td> 
<div align="center"></div>
</td>
</tr>
<tr> 
<td width="40%">antal brugere/&oslash;ko-net</td>
<td> 
<div align="center"><%=Application("users")%></div>
</td>
<td> 
<div align="center"><%=Application("eusers")%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("users").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("eusers").Value)%></div>
</td>
</tr>
</table>
<br>
<%
rsCount.Close()
%>
