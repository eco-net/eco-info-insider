<%
strConnect = "PROVIDER=MSDASQL;DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("db.mdb") & ";"
set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = strConnect
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 3
rs.Source="SELECT id,titel FROM tekst"
rs.Open()
set rsT = Server.CreateObject("ADODB.Recordset")
rsT.ActiveConnection = strConnect
rsT.CursorType = 0
rsT.CursorLocation = 2
rsT.LockType = 3
if request("id")<>"" then
rsT.Source="SELECT * FROM tekst Where id=" & request("id")
else
rsT.Source="SELECT * FROM tekst ORDER BY id desc"
end if
rsT.Open()
%>
<html>
<head>
<title>parkeringsterror</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#0044E7" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td>
      <table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td width="180">&nbsp;</td>
          <td width="360"> 
            <div align="center"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="5">parkeringsterror.dk</font></div>
          </td>
          <td width="180">&nbsp;</td>
        </tr>
        <tr> 
          <td width="180" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">Intro</font></td>
          <td width="360"> 
            <div align="center"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="350" height="266">
                <param name=movie value="hykler.swf">
                <param name=quality value=high>
                <param name="LOOP" value="false">
                <embed src="hykler.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="350" height="266" loop="false">
                </embed> 
              </object></div>
          </td>
          <td rowspan="2" valign="top" width="180"><font face="Verdana, Arial, Helvetica, sans-serif">L&aelig;s 
            indl&aelig;g: </font> <br>
<font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
            <% if not rs.EOF then
			while not rs.EOF %>
            <a href="index.asp?id=<%=rs("id")%>"><%=rs("titel")%></a><br>
            <%
			rs.MoveNext
			wend
			end if
			rs.Close()
			%>
            </font> <br>
            <a href="write.asp"><font face="Verdana, Arial, Helvetica, sans-serif">Skriv 
            indl&aelig;g</font></a></td>
        </tr>
        <tr> 
          <td width="180">&nbsp;</td>
          <td width="360"><font face="Verdana, Arial, Helvetica, sans-serif"> 
            <%if not rsT.EOF then %>
            <%=rsT("dato")%><br>
            <b><%=rsT("titel")%></b><br>
            indsendt af: <b><%=rsT("forfatter")%></b><br>
            <br>
            <%=rsT("tekst")%>
            <% end if %>
            </font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
<%rsT.Close()%>
