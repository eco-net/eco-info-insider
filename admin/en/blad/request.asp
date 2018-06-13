<!--#include virtual="/shared/sqlcheck.asp"-->
<% if request("update")<>"" then

kapID=request("kapid")
'response.write kapID & "<br>"
'response.write request("selSbj") & "<br>"

'3. Inserting Subjects
	theSQL=""
response.write request("selSbj").Count()
	FOR each Item in request("selSbj")
		theSQL="INSERT INTO enblad_kap_r_art (kapid,artid) VALUES (" & kapID & "," & item & ");" & chr(13)
	response.write theSQL
	'execCommand theSQL
	NEXT
	'response.end
	
end if
%>

</html>
