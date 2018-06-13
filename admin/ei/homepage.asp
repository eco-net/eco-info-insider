<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include file="backend.asp" -->

<%
SUB WriteSplashes(lang)
	DIM thefile,x,theOptions
	SET rs=Server.CreateObject("ADODB.recordset")
	rs.ActiveConnection=MM_data_STRING
	rs.Source="SELECT Count(*) as theCount FROM hw2_pageelements WHERE category=6 AND language='" & lang & "'"
	rs.Open()
	theCount=rs.Fields.Item("theCount").value
	rs.close
	IF theCount=2 THEN
		thefile=ReadFile("/backend/htmlsnippets/hp_splashes_2.txt")
	ELSE
		thefile=ReadFile("/backend/htmlsnippets/hp_splashes_3.txt")
	END IF
	
	'Replacing Splashes
	rs.Source="Select s.headline,s.content,s.imagepath,s.link,s.linkaction FROM hw2_pageelements pe LEFT JOIN hw2_splashes s ON pe.content_number=s.id WHERE pe.category=6 AND pe.language='" & lang & "' ORDER BY pe.number ASC"
	rs.Open()
	x=1
	WHILE NOT rs.EOF
		thetag="#Splash_Header" & CStr(x) & "#"
		thefile=replace(thefile,thetag,rs.Fields.Item("headline").value & "")
		thetag="#Splash_Text" & CStr(x) & "#"
		thefile=replace(thefile,thetag,replace(rs.Fields.Item("content").value & "",chr(13),"<br>") & "")
		thetag="#Splash_Image" & CStr(x) & "#"
		thefile=replace(thefile,thetag,rs.Fields.Item("imagepath").value & "")
		thetag="#Splash_Link" & CStr(x) & "#"
		thefile=replace(thefile,thetag,rs.Fields.Item("link").value & "")
		thetag="#Splash_LinkAction" & CStr(x) & "#"
		thefile=replace(thefile,thetag,rs.Fields.Item("linkaction").value & "")
		rs.movenext
		x=x+1
	WEND
	rs.close()

	'Replacing News
	rs.Source="Select n.id,n.headline,n.content FROM hw2_pageelements pe LEFT JOIN hw2_news n ON pe.content_number=n.id WHERE pe.category=10 AND pe.language='" & lang & "' ORDER BY pe.number ASC"
	rs.Open()
	x=1
	WHILE NOT rs.EOF
		thetag="#News_Header" & CStr(x) & "#"
		thefile=replace(thefile,thetag,rs.Fields.Item("headline").value & "")
		thetag="#News_Text" & CStr(x) & "#"
		thefile=replace(thefile,thetag,replace(rs.Fields.Item("content").value & "",chr(13),"<br>") & "")
		thetag="#News_Link" & CStr(x) & "#"
		thefile=replace(thefile,thetag,"/about/news_detail.asp?id=" & rs.Fields.Item("id").value)
		rs.movenext
		x=x+1
	WEND
	rs.close()
	'Replacing remaining news-tags with empty values
	IF x<=3 THEN
		WHILE x<4
			thetag="#News_Header" & CStr(x) & "#"
			thefile=replace(thefile,thetag,"")
			thetag="#News_Text" & CStr(x) & "#"
			thefile=replace(thefile,thetag,"")
			thetag="#News_Link" & CStr(x) & "#"
			thefile=replace(thefile,thetag,"")
			x=x+1		
		WEND
	END IF

	'Replacing dealer options
'	rs.Source="Select * FROM hw2_pageelements WHERE category=6 AND language='" & lang & "' ORDER BY number ASC"
'	rs.Open()
	theOptions="<option value=""0"" selected>&lt;Produkttype&gt;</option>"
'	WHILE NOT rs.EOF
'		theOptions=theOptions & chr(13) & "<option value=""" & rs.fields.item("id").value & """>" & rs.Fields.Item("name").value & "</option>" 
'	WEND
	thefile=replace(thefile,"#Dealer_Options#",theOptions)
'	rs.close()
	rs=""
	WriteFile thefile,"/" & lang & "/home/inc_splashes.txt"
END SUB

Sub WriteBa(lang)
	SET rs=Server.CreateObject("ADODB.recordset")
	rs.ActiveConnection=MM_data_STRING
	rs.Source="Select count(*) as thecount FROM hw2_pageelements pe WHERE pe.category=7 AND pe.language='" & lang & "'"
	rs.Open()
	theCount=rs.Fields.Item("thecount").value
	rs.Close()

	thefile=ReadFile("/backend/htmlsnippets/hp_branding.txt")
	rs.Source="Select b.id,b.headline,b.content,b.imagepath,b.link,b.linkaction FROM hw2_pageelements pe LEFT JOIN hw2_branding b ON pe.content_number=b.id WHERE pe.category=7 AND pe.language='" & lang & "' ORDER BY pe.number ASC"
	rs.Open()
	i=1
	WHILE NOT rs.EOF
		theText=replace(thefile,"#Branding_Image#",rs.Fields.Item("imagepath").value & "")
		theText=replace(theText,"#Branding_Header#",rs.Fields.Item("headline").value & "")
		theText=replace(theText,"#Branding_Text#",rs.Fields.Item("content").value & "")
		theText=replace(theText,"#Branding_Link#",rs.Fields.Item("link").value & "")
		theText=replace(theText,"#Branding_LinkAction#",rs.Fields.Item("linkaction").value & "")
		IF theCount=1 THEN
			WriteFile theText,"/" & lang & "/home/inc_branding.txt"
		ELSE
			WriteFile theText,"/" & lang & "/home/inc_branding_" & CStr(i) & ".txt"
		END IF
		i=i+1
		rs.movenext
	WEND
	IF theCount>1 THEN
		thefile=ReadFile("/backend/htmlsnippets/hp_branding_random.asp")
		thefile=replace(thefile,"#count#",CStr(thecount))
		WriteFile thefile,"/" & lang & "/home/inc_branding.txt"
	END IF
END SUB
%>