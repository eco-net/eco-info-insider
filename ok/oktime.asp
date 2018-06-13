<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<body>
<form id="form1" name="form1" method="post" action="">
  <select name="oktime" class="formInputobject">
    <script src="/log/optionfiles/oktime_options2.js"></script>
  </select>
</form>
</body>
</html>
<script language="javascript">
idag=new Date();
m=idag.getMonth();
y=idag.getYear();

for(var i=0;i<12;i++){
writeMonth(m,y);
m++;
if (m>11){
m=0;
y++;
}
}
</script>