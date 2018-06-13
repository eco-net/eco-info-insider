<%
IF LEN(request.cookies("eiuserid"))=0 THEN response.redirect "index.asp"
%>