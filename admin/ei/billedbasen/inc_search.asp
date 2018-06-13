
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
set rsCat = Server.CreateObject("ADODB.Recordset")
rsCat.ActiveConnection = MM_ecoinfo_STRING
rsCat.Source = "SELECT *  FROM image_cat_maindata  ORDER BY name"
rsCat.CursorType = 0
rsCat.CursorLocation = 2
rsCat.LockType = 3
rsCat.Open()
rsCat_numRows = 0
%>
<%
set rsSize = Server.CreateObject("ADODB.Recordset")
rsSize.ActiveConnection = MM_ecoinfo_STRING
rsSize.Source = "SELECT *  FROM image_size"
rsSize.CursorType = 0
rsSize.CursorLocation = 2
rsSize.LockType = 3
rsSize.Open()
rsSize_numRows = 0
%>
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
<table width="100%" border="0" cellspacing="0" cellpadding="5">
<tr> 
<td> 
<form name="form1" method="post" action="index.asp">
<p class="listheader" align="center">S&oslash;g og du skal finde</p>
<p align="center"> 
<select name="cat" class="formInputobject">
<option value="" selected>Vælg Kategori</option>
<%
While (NOT rsCat.EOF)
%>
<option value="<%=(rsCat.Fields.Item("id").Value)%>" ><%=(rsCat.Fields.Item("name").Value)%></option>
<%
  rsCat.MoveNext()
Wend
If (rsCat.CursorType > 0) Then
  rsCat.MoveFirst
Else
  rsCat.Requery
End If
%>
</select>
</p>
<p align="center"> 
<select name="sizes" class="formInputobject">
<option value="" selected>Vælg Størrelse</option>
<%
While (NOT rsSize.EOF)
%>
<option value="<%=(rsSize.Fields.Item("id").Value)%>" ><%=(rsSize.Fields.Item("name").Value)%></option>
<%
  rsSize.MoveNext()
Wend
If (rsSize.CursorType > 0) Then
  rsSize.MoveFirst
Else
  rsSize.Requery
End If
%>
</select>
<br>
<span class="listheader">id</span><br>
<input type="text" name="img_id" class="formInputobject">
<br>
<span class="listheader">s&oslash;geord</span> <br>
<input type="text" name="searchtext" class="formInputobject" value="">
</p>
<p align="center">
<input type="hidden" name="search" value="1">
<input type="submit" name="Submit" value="S&oslash;g" class="formSubmitbutton">
</p>
</form>
</td>
</tr>
</table>
<%
rsCat.Close()
%>
<%
rsSize.Close()
%>
