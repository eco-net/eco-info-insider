<% response.write request("elm1") %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<script type="text/javascript" src="tinymce/jscripts/tiny_mce/tiny_mce.js"></script>
 <link rel="stylesheet" type="text/css" href="/tiny/woody01/style.css" />
 <!--#include virtual="/shared/common.asp"-->
 <%'tekst=readFile("/tiny/woody01/index.html")%>
 <%tekst=request("elm1")%>
</head>

<body>
<script language="javascript">

tinyMCE.init({
	mode : "textareas",
	theme : "advanced"
});

</script>
<form action="index.asp" method="post" name="form1" id="form1">
  <textarea id="elm1" name="elm1" rows="50" cols="80" style="width: 80%">
	<%=tekst%>
	</textarea>
  <p>
    <label>
    <input type="submit" name="Submit" value="Submit" />
    </label>
</p>
</form>
</body>
</html>
