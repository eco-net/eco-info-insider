
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
set rsCount = Server.CreateObject("ADODB.Recordset")
rsCount.ActiveConnection = MM_ecoinfo_STRING
rsCount.Source = "SELECT *  FROM eisys_bu_count"
rsCount.CursorType = 0
rsCount.CursorLocation = 2
rsCount.LockType = 3
rsCount.Open()
rsCount_numRows = 0
%>
<%


%>
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
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
<div align="center" class="contentHeader1">BU</div>
</td>
<td class="listheader" colspan="2"> 
<div align="center">fra 9/12-02</div>
</td>
</tr>
<tr> 
<td width="40%">&nbsp;</td>
<td width="15%">&nbsp;</td>
<td width="15%">&nbsp;</td>
<td width="15%"> 
<div align="center"><font color="#FF0000">bruger</font></div>
</td>
<td> 
<div align="center"><font color="#0000FF">&oslash;ko-net</font></div>
</td>
</tr>
<tr> 
<td width="40%">home/index</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("homeindex").Value)%></div>
</td>
<td width="15%"> 
<div align="center"><%=(rsCount.Fields.Item("ehomeindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">hvad/index</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("hvadindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("ehvadindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">aktuelt/index</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("aktueltindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("eaktueltindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">kalender/list</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("oklist").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("eoklist").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">netværksted/index</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("netvindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("enetvindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">uddannelse/index</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("uddindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("euddindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">hvad/baggrund/index</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("hvadbrindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("ehvadbrindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">netv&aelig;rksted/idebank</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("netidebankindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("enetidebankindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">netv&aelig;rksted/idekuv&oslash;se</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("netidekuvindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("enetidekuvindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">uddannelse/udd &amp; kurser</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("uddorguddkur").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("euddorguddkur").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">links/uddindex</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("hvadlinkuddindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("ehvadlinkuddindex").Value)%></div>
</td>
</tr>

<tr> 
<td class="listheader" width="40%">Sessions</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"></div>
</td>
<td> 
<div align="center"></div>
</td>
</tr>
<tr> 
<td width="40%">antal brugere/&oslash;ko-net</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
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
