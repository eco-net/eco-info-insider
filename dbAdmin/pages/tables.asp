 <!--include virtual="/dbadmin/connections/voresDebat.asp" -->
  <!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/dbadmin/pages/act_struct_tables.asp" -->

<link rel="stylesheet" href="/dbAdmin/shared/styles.css" type="text/css">
<%
Dim rsTables__ColParam
rsTables__ColParam = "%"
if (request("thename") <> "") then rsTables__ColParam = request("thename")
%>
<%
set rsTables = Server.CreateObject("ADODB.Recordset")
'rsTables.ActiveConnection = MM_voresDebat_STRING
rsTables.ActiveConnection = MM_ecoinfo_STRING
rsTables.Source = "SELECT *  FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME LIKE '%" + Replace(rsTables__ColParam, "'", "''") + "%'"
rsTables.CursorType = 0
rsTables.CursorLocation = 2
rsTables.LockType = 3
rsTables.Open()

rsTables_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsTables_numRows = rsTables_numRows + Repeat1__numRows
%>
<script src="/dbAdmin/shared/list2.js"></script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr class="listlabelbg"> 
<td width="40%"><b> Tabel struktur </b></td>
<td><b>Data</b></td>
<td>&nbsp;</td>
</tr>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsTables.EOF)) 
%>
<tr> 
<td class="listtext"><a href="javascript:openColumns('<%=(rsTables.Fields.Item("TABLE_NAME").Value)%>')"><%=(rsTables.Fields.Item("TABLE_NAME").Value)%></a></td>
<td class="listtext"><a href="javascript:showData('<%=(rsTables.Fields.Item("TABLE_NAME").Value)%>')">Vis Data</a>&nbsp;</td>
<td class="listtext">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td>
</tr>
<tr>
<td colspan="3" height="2"><img src="/dbadmin/images/spacer.gif" height="2" width="1" border="0"></td>
</tr>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsTables.MoveNext()
Wend
%>
</table>
<%
rsTables.Close()
%>