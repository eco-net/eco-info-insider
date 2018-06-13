
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
set rsCount = Server.CreateObject("ADODB.Recordset")
rsCount.ActiveConnection = MM_ecoinfo_STRING
rsCount.Source = "SELECT *  FROM eisys_econet_count"
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
<div align="center" class="contentHeader1">&Oslash;ko-net</div>
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
<td width="40%">blad/index</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("bladindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("ebladindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">udgivelser/index</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("udgivindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("eudgivindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">seminarer/list</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("seminarindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("eseminarindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">om/index</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("omindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("eomindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">english/index</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("engindex").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("eengindex").Value)%></div>
</td>
</tr>
<tr> 
<td width="40%">&aring;rsm&oslash;de 2003</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("semaarsmode").Value)%></div>
</td>
<td> 
<div align="center"><%=(rsCount.Fields.Item("esemaarsmode").Value)%></div>
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
