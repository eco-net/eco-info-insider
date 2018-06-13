<%
DIM theText
IF CInt(request("done"))=1 THEN
	theText="Nyt Tema er valgt"
ELSEIF CInt(request("done"))=2 THEN
	theText="Aktuelt tema er opdateret på Øko-info"
ELSEIF CInt(request("done"))<5 THEN
	theText="You have succesfully removed the Splash."
ELSE
	theText="You have succesfully updated Splashes on Homepage."
END IF
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/backend/backend.css" type="text/css">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="functions.asp">
<table border="0" cellspacing="0" cellpadding="0" class="functionstable" align="center">
<tr> 

<td class="functionsheader">&nbsp;</td>
</tr>
<tr> 
<td> 
<div align="center"></div>
</td>
<td align="center"> 
<table width="350" border="0" cellspacing="0" cellpadding="0" class="alertbox">
<tr class="functionstable"> 
<td class="contentHeader1" align="center">Bekr&aelig;ftelse</td>
</tr>
<tr class="functionstable"> 
<td align="center" class="alerttext"> 
<p><%=theText%>
<br>
<br>
<input type="submit" name="Submit2" value="OK" class="formSubmitbutton">
</p>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</body>
</html>


