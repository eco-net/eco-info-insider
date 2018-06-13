<!--#include virtual="/shared/common.asp"-->
<%
theSQL="DELETE FROM eisys_fmpmail"
		execCommand theSQL
	filename=Server.MapPath("mailadresser.txt")
	SET fs = CreateObject("Scripting.FileSystemObject")
		Set ts=fs.OpenTextFile(filename)
		i=0
		while not ts.AtEndOfStream
		i=i+1
		email=ts.ReadLine()
		theSQL="INSERT INTO eisys_fmpmail (mail) VALUES ('" & email & "')"
		execCommand theSQL
		response.write i & " - " & theSQL & "<br>"
		wend
%>


