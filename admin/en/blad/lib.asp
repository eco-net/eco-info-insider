<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/common.asp"-->
<% 
Function doSave()
if request("update")<>"" then
kapID=request("kapid")

' Inserting Art and Kap
	
	

		theSQL="DELETE FROM enblad_kap_r_art WHERE kap_id=" & kapID
		execCommand theSQL
		theSQL=""
			
	FOR each Item in request("selSbj")
		theSQL="INSERT INTO enblad_kap_r_art (kap_id,art_id) VALUES (" & kapID & "," & item & ");" & chr(13)
		execCommand theSQL
		theSQL=""
	NEXT	
	
end if
end function
%>