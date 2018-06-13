
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/common.asp" -->
<%Dim idag
idag=CStr(FormatDateTime(Date,2))

%>


<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT o.id, o.name, o.fullname, s.username,s.page,s.date,s.app_id,s.sek,s.app_title  FROM eiorg_maindata o INNER JOIN eisys_insideruser_stat s ON  o.id=s.orgid WHERE s.dato LIKE'" & idag & "' ORDER BY date desc"
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
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
<table align="center" width="380" border="0" cellspacing="0" cellpadding="5" background="/dgs/graphics/sub_index_header_dgs_bkgrd.gif" name="subIndexHeader">
<tr> 
<td valign="top"> 
<table width="95%" border="0" cellspacing="0" cellpadding="0" background="/shared/graphics/spacer.gif">
<tr> 
<td class="contentHeader1"> 
<% If Not rsPageData.EOF Or Not rsPageData.BOF Then %>
Disse brugere har idag rettet deres data: 
<% else %>
Der er ingen brugere der har rettet deres data idag 
<% End If ' end Not rsPageData.EOF Or NOT rsPageData.BOF %>
</td>
</tr>
<tr> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr> 
<td valign="top" class="sitetag"> 
<table width="100%" border="0" cellspacing="0" cellpadding="5">
<% If Not rsPageData.EOF Or Not rsPageData.BOF Then %>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsPageData.EOF)) 
%>
<tr> 
<td width="30%"><a href="/dgs/edit.asp?id=<%=(rsPageData.Fields.Item("id").Value)%>" title="<%=(rsPageData.Fields.Item("username").Value)%>"> 
<%if(rsPageData.Fields.Item("name").Value)<>"" then %>
<%=(rsPageData.Fields.Item("name").Value)%> 
<%else %>
<%=(rsPageData.Fields.Item("fullname").Value)%> 
<% end if %>
</a></td>
<td width="10%"> 
<%if(rsPageData.Fields.Item("page").Value)=1 then %>
oprettede 
<%else %>
redigerede 
<%end if %>
</td>
<td width="10%"><%=rsPageData("sek")%></td>
<td width="35%"><a href="/<%=(rsPageData.Fields.Item("sek").Value)%>/edit.asp?id=<%=(rsPageData.Fields.Item("app_id").Value)%>"><%=(rsPageData.Fields.Item("app_title").Value)%></a></td>
<td width="15%" class="sitetag"><%=Datepart("h",(rsPageData.Fields.Item("date").Value))%>:<%=right("0" & CSTR(datepart("n",rsPageData.Fields.Item("date").Value)),2) %> </td>
</tr>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsPageData.MoveNext()
Wend
%>
<% End If ' end Not rsPageData.EOF Or NOT rsPageData.BOF %>
</table>


</td>
</tr>

</table>
<div align="center"><br>
<a href="/home/stathistory.asp">Tidligere brugere</a></div>
</td>
</tr>
</table>
<%
rsPageData.Close()
%>
<!--#include file="visitstat.asp"-->
<!--#include file="ecoinfostat.asp"-->
<!--#include file="bustat.asp"-->
<!--#include file="econetstat.asp"-->
