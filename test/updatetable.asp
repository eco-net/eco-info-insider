<!--#include virtual="/shared/common.asp"-->
<%
'theSQL="ALTER TABLE bu_Artikel ADD miljoinfo integer DEFAULT 0 "
'theSQL="UPDATE bu_Artikel SET miljoinfo=0"
'theSQL="CREATE TABLE eisys_returnmail (id integer identity, mail varchar(100), PRIMARY KEY (id));"
'theSQL="CREATE TABLE ubu_news_maindata (id integer identity, title varchar(100), createdate datetime DEFAULT getdate(), shortdescription varchar(1500), description varchar(8000), right_col varchar(1000), PRIMARY KEY (id));"
'response.write theSQL
'theSQL="DELETE FROM eisys_testmail WHERE id=15"
'theSQL="UPDATE eisys_insiderusers SET password='borgaard' WHERE id=14605"
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


