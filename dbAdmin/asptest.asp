<%
SUB MakeFile(theSQL,theConn,StartText,EndText,ModelText,Filename)
	DIM theRS, Format_Sel, fso, ts,theData,rowCount, colCount,i,ii, thisRow, theFile
	SET theRS= Server.CreateObject("ADODB.Recordset")
	theRS.ActiveConnection = theConn
	theRS.Source = theSQL
	theRS.Open()
'	WHILE NOT theRS.EOF
'		response.write ("her<br>")
'		theRS.movenext
'	WEND
	theData=theRS.GetRows
	theRS.Close()
	colCount=uBound(theData,2)
	rowCount=uBound(theData,1)
	theFile=StartText
	FOR i=0 TO colCount
		thisRow=ModelText
		FOR ii=0 TO rowCount
			thisRow=replace(thisRow,"#" & ii & "#",theData(ii,i) & " ")
		NEXT
		theFile=theFile & thisrow
	NEXT
	theFile=theFile & EndText
	WriteFile theFile, Filename
END SUB

SUB WriteFile(theText,filename)
	DIM fso,theFile
	filename=server.MapPath(filename)
	SET fso = createobject("SoftArtisans.FileManager")
	IF NOT fso.FileExists(filename) THEN 
		Set ts=fso.CreateTextFile(filename)
	END IF
	SET theFile = fso.OpenTextFile(filename,2)
	theFile.Write(theText)
	fso=""
	theFile=""
END SUB

%>