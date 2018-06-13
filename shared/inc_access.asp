<%
IF LEN(request.cookies("eiuserid"))=0 AND LEN(request.cookies("eiorgid"))=0 THEN response.redirect "index.asp"
%>