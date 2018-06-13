<!--#include virtual="/Connections/ecoinfo.asp" -->
<%

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_ecoinfo_STRING
rs.Source = "SELECT * FROM stat_counter ORDER BY date desc"
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 3
rs.Open()
rs_numRows = 0
themonth=DatePart("m", rs("date")) 'hvilken måned starter vi
theYear=DatePart("yyyy", rs("date"))
infocount=0
netcount=0
insidercount=0
allinfo=0
allnet=0
allinsider=0
allmonth=1
gennemsnit="<br><center><H3>Sessions i snit pr. måned</H3></center><table width=95% border=0 align=center><tr><td><b>måneder</b></td><td align=center>&nbsp;</td><td align=center><b>Øko-info</b></td><td align=center><b>Øko-net</b></td><td align=center><b>Insider</b></td></tr>"
thetable="<center><H3>Sessions pr. måned</H3></center><table width=95% border=1 align=center><tr><td><b>måned</td><td align=center><b>år</td></b><td align=center><b>Øko-info</b></td><td align=center><b>Øko-net</b></td><td align=center><b>Insider</b></td></tr>"
countSessions
response.write gennemsnit
response.write thetable


Function countSessions()
while not rs.EOF 
	if themonth<>DatePart("m", rs("date")) then
	allmonth=allmonth + 1
	thetable=thetable & "<tr><td>" & mnavn(themonth) & "</td><td align=center>" & theYear & "</td><td align=center>" & infocount & "</td><td align=center>" & netcount & "</td><td align=center>" & insidercount & "</td></tr>"
		themonth=DatePart("m", rs("date"))
		theYear=DatePart("yyyy", rs("date"))
		infocount=0
		netcount=0
		insidercount=0
	end if
	infocount=infocount + CInt(rs("ecoinfo"))
	netcount=netcount + CInt(rs("econet"))
	insidercount=insidercount + CInt(rs("insider"))
	allinfo=allinfo + CInt(rs("ecoinfo"))
	allnet=allnet + CInt(rs("econet"))
	allinsider=allinsider + CInt(rs("insider"))
rs.MoveNext
wend
gennemsnit=gennemsnit & "<tr><td>" & allmonth & "</td><td align=center>" & "&nbsp;</td><td align=center>" & CInt(allinfo/allmonth) & "</td><td align=center>" & CInt(allnet/allmonth) & "</td><td align=center>" & CInt(allinsider/allmonth) & "</td></tr></table>"
thetable=thetable & "</table>"
end function

Function mnavn(m)

   Select Case m
      Case 1   mnavn="Januar"
      Case 2   mnavn="Februar"
	  Case 3   mnavn="Marts"
	  Case 4   mnavn="April"
	  Case 5   mnavn="Maj"
	  Case 6   mnavn="Juni"
	  Case 7   mnavn="Juli"
	  Case 8   mnavn="August"
      Case 9   mnavn="September"
	  Case 10  mnavn="Oktober"
	  Case 11  mnavn="November"
	  Case 12  mnavn="December"
   End Select

End Function


rs.close


%>