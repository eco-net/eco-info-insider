<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp"-->
<%
if request("update")<>"" then
theSQL="INSERT INTO eisys_testmail (mail) Values('" & request("mail") & "')"
execCommand theSQL
end if
	set rs=server.createobject("ADODB.recordset")
	rs.Activeconnection=MM_ecoinfo_STRING
	rs.source="SELECT mail FROM eisys_testmail"
	rs.open()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <table width="100%" border="0" cellpadding="5">
    <tr>
      <td>ny mail </td>
      <td><label>
        <input name="mail" type="text" id="mail" />
      </label></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><label>
        <input type="submit" name="Submit" value="Submit" />
      </label>
      <input name="update" type="hidden" id="update" value="1" /></td>
    </tr>
  </table>
</form>
<br />

<table width="100%" border="0" cellpadding="5">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <%WHILE NOT rs.EOF %>
		
  <tr>
    <td><%=rs("mail")%></td>
  </tr>
  <%rs.MoveNext
  Wend
  rs.Close
  %>
</table>
</body>
</html>
