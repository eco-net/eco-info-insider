<%
FUNCTION ReadFile(filename)
	DIM fs,ts,theText
	theText="file created"
	filename=Server.MapPath(filename)
	SET fs = CreateObject("Scripting.FileSystemObject")
	IF fs.FileExists(filename)=false THEN
		SET ts = fs.CreateTextFile(filename)
		ts.Write(theText)
	ELSE		
		Set ts=fs.OpenTextFile(filename)
		theText=ts.ReadAll()
	END IF
	fs=""
	ts=""
	ReadFile=theText
END FUNCTION
SUB WriteFile(theText,filename)
	DIM fs,ts
	filename=server.MapPath(filename)
	SET fs = createobject("Scripting.FileSystemObject")
	Set ts=fs.CreateTextFile(filename,true)
	ts.Write(theText)
	fs=""
	ts=""
END SUB
%>