<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/inc_adm_access.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = MM_ecoinfo_STRING
  MM_editTable = "eiafstem_ning"
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "forslag|value|descr|value|emne_id|value"
  MM_columnsStr = "forslag|',none,''|descr|',none,''|emne_id|none,none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
set rsAfstem = Server.CreateObject("ADODB.Recordset")
rsAfstem.ActiveConnection = MM_ecoinfo_STRING
rsAfstem.Source = "SELECT *  FROM eiafstem_ning"
rsAfstem.CursorType = 0
rsAfstem.CursorLocation = 2
rsAfstem.LockType = 3
rsAfstem.Open()
rsAfstem_numRows = 0
%>
<%
Dim rsAfstemEmne__theID
rsAfstemEmne__theID = "0"
if (request("id") <> "") then rsAfstemEmne__theID = request("id")
%>
<%
set rsAfstemEmne = Server.CreateObject("ADODB.Recordset")
rsAfstemEmne.ActiveConnection = MM_ecoinfo_STRING
rsAfstemEmne.Source = "SELECT *  FROM eiafstem_emne  WHERE id=" + Replace(rsAfstemEmne__theID, "'", "''") + ""
rsAfstemEmne.CursorType = 0
rsAfstemEmne.CursorLocation = 2
rsAfstemEmne.LockType = 3
rsAfstemEmne.Open()
rsAfstemEmne_numRows = 0
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="center"><span class="contentHeader2">Start Afstemning</span> <br>
<%=(rsAfstemEmne.Fields.Item("name").Value)%></p>
<p align="center">Kom med et forslag og afstemningen er igang p&aring; forsiden.</p>
<form name="form1" method="POST" action="<%=MM_editAction%>">
<p>&nbsp; </p>
<p align="center">Forslag<br>
<input type="text" name="forslag" class="formInputobject">
<br>
Beskrivelse<br>
<input type="text" name="descr" class="formInputobjectLong">
</p>
<p align="center"> 
<input type="hidden" name="emne_id" value=<%=request("id")%>>
<input type="submit" name="Submit" value="gem" class="formSubmitbutton">
</p>
<input type="hidden" name="MM_insert" value="true">
</form>
<p>&nbsp;</p>
<p><br>
</p>
</body>
</html>
<%
rsAfstem.Close()
%>
<%
rsAfstemEmne.Close()
%>

