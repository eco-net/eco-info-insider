<!--include virtual="/shared/sqlcheck.asp"-->
<%
FUNCTION GetParams(removeparams)
	' create the list of parameters which should not be maintained
	MM_removeList = "&d=" & removeparams
	' add the URL parameters to the MM_keepURL string
	For Each Item In Request.QueryString
	  NextItem = "&" & Item & "="
	  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
		MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
	  End If
	Next
	response.write "<input type=""hidden"" name=""params"" value=""" & MM_keepURL & """>"
END FUNCTION
%>
