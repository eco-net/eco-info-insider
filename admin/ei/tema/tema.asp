<!--#include virtual="shared/common.asp"-->
<%
theSQL="INSERT INTO en_mainpage (title,descr,img,link,linktext)" &_
" VALUES ('Titel','Desription',0,'Link','Linktext')"
'theSQL="UPDATE en_mainpage SET title='Titel',descr='description,img=0,link='link',linktext='linktext'"
execCommand theSQL
response.write "OK"
%>