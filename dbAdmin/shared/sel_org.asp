
<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/ecopages_adm/Connections/ecopages.asp" -->
<!--#include virtual="/ecopages_adm/shared/restrictAccess.asp" -->
<%
Dim RecCount
set rsRecCount = Server.CreateObject("ADODB.Recordset")
rsRecCount.ActiveConnection = MM_ecopages_STRING
rsRecCount.Source = "SELECT Count(*) as recCount FROM gd_organisations WHERE organisation_name LIKE'" + Replace(request("orgname"), "'", "''") + "%'"
rsRecCount.CursorType = 0
rsRecCount.CursorLocation = 2
rsRecCount.LockType = 3
rsRecCount.Open()
RecCount=rsRecCount.Fields.Item("recCount").value
rsRecCount.Close()
%>
<%
Dim rsPageList__MM_ColParam
rsPageList__MM_ColParam = "%"
if (request("orgname") <> "") then rsPageList__MM_ColParam = request("orgname")
%>

<%
set rsPageList = Server.CreateObject("ADODB.Recordset")
rsPageList.ActiveConnection = MM_ecopages_STRING
rsPageList.Source = "SELECT organisation_id, organisation_name, zip, city  FROM gd_organisations  WHERE organisation_name LIKE'" + Replace(rsPageList__MM_ColParam, "'", "''") + "%'  ORDER BY organisation_name"
rsPageList.CursorType = 0
rsPageList.CursorLocation = 2
rsPageList.LockType = 3
rsPageList.Open()
rsPageList_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 100
Dim Repeat1__index
Repeat1__index = 0
rsPageList_numRows = rsPageList_numRows + Repeat1__numRows
%>


<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsPageList_total = rsPageList.RecordCount
rsPageList_total = RecCount

' set the number of rows displayed on this page
If (rsPageList_numRows < 0) Then
  rsPageList_numRows = rsPageList_total
Elseif (rsPageList_numRows = 0) Then
  rsPageList_numRows = 1
End If

' set the first and last displayed record
rsPageList_first = 1
rsPageList_last  = rsPageList_first + rsPageList_numRows - 1

' if we have the correct record count, check the other stats
If (rsPageList_total <> -1) Then
  If (rsPageList_first > rsPageList_total) Then rsPageList_first = rsPageList_total
  If (rsPageList_last > rsPageList_total) Then rsPageList_last = rsPageList_total
  If (rsPageList_numRows > rsPageList_total) Then rsPageList_numRows = rsPageList_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsPageList_total = -1) Then

  ' count the total records by iterating through the recordset
  rsPageList_total=0
  While (Not rsPageList.EOF)
    rsPageList_total = rsPageList_total + 1
    rsPageList.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsPageList.CursorType > 0) Then
    rsPageList.MoveFirst
  Else
    rsPageList.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsPageList_numRows < 0 Or rsPageList_numRows > rsPageList_total) Then
    rsPageList_numRows = rsPageList_total
  End If

  ' set the first and last displayed record
  rsPageList_first = 1
  rsPageList_last = rsPageList_first + rsPageList_numRows - 1
  If (rsPageList_first > rsPageList_total) Then rsPageList_first = rsPageList_total
  If (rsPageList_last > rsPageList_total) Then rsPageList_last = rsPageList_total

End If
%>


<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsPageList
MM_rsCount   = rsPageList_total
MM_size      = rsPageList_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>

<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>

<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>

<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsPageList_first = MM_offset + 1
rsPageList_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsPageList_first > MM_rsCount) Then rsPageList_first = MM_rsCount
  If (rsPageList_last > MM_rsCount) Then rsPageList_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>

<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>

<%
' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
%>

<html><!-- #BeginTemplate "/Templates/list.dwt" --><!-- DW6 -->
<head>
<!-- #BeginEditable "doctitle" -->
<title>Untitled Document</title>
<script language="JavaScript">
idfield='<%=request("idfield")%>'
namefield='<%=request("namefield")%>'

function selectorg(name,id) {
	eval('opener.document.theForm.' + idfield + '.value=' + id)
	eval("opener.document.theForm." + namefield + ".value='" + name + "'")
	window.close()
}
</script>
<!--#include virtual="/ecopages_adm/shared/closesubwin.asp" -->
<!-- #EndEditable -->
<script src="/dbadmin/shared/list.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/dbadmin/shared/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="3" marginwidth="0" marginheight="3">
<table width="750" border="0" cellspacing="0" cellpadding="0" align="center"><tr>
<td valign="top">
<!-- START Menu -->
<!-- #BeginEditable "menu" -->

<!-- #EndEditable -->
<!-- END Menu -->

<!-- START ContentHeader -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td background="/dbadmin/images/layout/page_l.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="1"></td>
<td class="tableheader"  width="98%"> <!-- #BeginEditable "overskrift" -->&nbsp;Fundne organisationer&nbsp;&nbsp;<span class="listtext"><br>
&nbsp;(<%=(rsPageList_first)%>-<%=(rsPageList_last)%> af <%=(rsPageList_total)%></span>)
<!-- #EndEditable --></td>
<td align="right"> <!-- #BeginEditable "navigation" -->
<% If MM_offset = 0 Then %>
<span class="tableheaderlinkdisabled">Forrige side</span>
<% End If ' end MM_offset = 0 %>

<% If MM_offset <> 0 Then %>
<a href="<%=MM_movePrev%>">Forrige side</a>
<% End If ' end MM_offset <> 0 %>
&nbsp;|&nbsp;
<% If MM_atTotal Then %>
<span class="tableheaderlinkdisabled">N&aelig;ste side</span>
<% End If ' end MM_atTotal %>

<% If Not MM_atTotal Then %>
<a href="<%=MM_moveNext%>">N&aelig;ste side</a>
<% End If ' end Not MM_atTotal %>
&nbsp;
<!-- #EndEditable --> 
</td>
<td background="/dbadmin/images/layout/page_r.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="1"></td>
</tr>
<tr>
<td background="/dbadmin/images/layout/page_l.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
<td><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
<td><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
<td background="/dbadmin/images/layout/page_r.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
</tr>
</table>
<!-- END ContentHeader -->

<!-- START Content -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td background="/dbadmin/images/layout/page_l.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
<td width="99%">
<!-- #BeginEditable "indhold" -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td>

<% If Not rsPageList.EOF Or Not rsPageList.BOF Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">

<% 
While ((Repeat1__numRows <> 0) AND (NOT rsPageList.EOF)) 
%>
<tr>
<td class="listtext"><span class="formlabelsleft"><%=(rsPageList.Fields.Item("organisation_name").Value)%></span><br>
<%=(rsPageList.Fields.Item("zip").Value)%>&nbsp;<%=(rsPageList.Fields.Item("city").Value)%></td>
<td class="listtext" align="right"><a href="javascript:selectorg(<%="'" & (rsPageList.Fields.Item("organisation_name").Value) & "','" & (rsPageList.Fields.Item("organisation_id").Value) & "'" %>)">V&aelig;lg</a><br>
&nbsp;
</td>
</tr>
<tr>
<td><img src="/dbadmin/images/listline.gif" height="3" width="100%" align="center"></td>
<td><img src="/dbadmin/images/listline.gif" height="3" width="100%" align="center"></td>
</tr>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsPageList.MoveNext()
Wend
%>
<tr>
<td class="tableheader" width="50%"></td>
<td class="tableheaderlink" width="50%">
<% If MM_offset = 0 Then %>
<span class="tableheaderlinkdisabled">Forrige side</span>
<% End If ' end MM_offset = 0 %>

<% If MM_offset <> 0 Then %>
<a href="<%=MM_movePrev%>">Forrige side</a>
<% End If ' end MM_offset <> 0 %>
&nbsp;|&nbsp;
<% If MM_atTotal Then %>
<span class="tableheaderlinkdisabled">N&aelig;ste side</span>
<% End If ' end MM_atTotal %>

<% If Not MM_atTotal Then %>
<a href="<%=MM_moveNext%>">N&aelig;ste side</a>
<% End If ' end Not MM_atTotal %>
&nbsp;
</td>
</tr>
</table>
<% End If ' end Not rsPageList.EOF Or NOT rsPageList.BOF %>

<% If rsPageList.EOF And rsPageList.BOF Then %>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td class="formlabelsleft">Der er ingen organisationer hvis navn begynder med <%=request("orgname")%><br>
<a href="javascript:window.close();" class="listtext">Luk vindue</a></td>
</tr>
</table>

<% End If ' end rsPageList.EOF And rsPageList.BOF %>


</td>
</tr>

</table>
<!-- #EndEditable -->
</td>
<td background="/dbadmin/images/layout/page_r.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
</tr>
</table>
<!-- END content -->

<!-- START Footer -->
<table border="0" width="100%" cellpadding="0" cellspacing="0">
<tr>
<td><img src="/dbadmin/images/layout/page_bl.gif" border="0" width="12" height="12"></td>
<td background="/dbadmin/images/layout/page_b.gif" width="99%"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="1" height="12"></td>
<td><img src="/dbadmin/images/layout/page_br.gif" border="0" width="12" height="12"></td>
</tr>
</table>
</td></tr></table>
</body>
<!-- #EndTemplate --></html>
<%
rsPageList.Close()
%>




