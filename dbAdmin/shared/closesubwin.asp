<%if (CStr(request("closesub"))="1") then
	response.write("<script language=""JavaScript"">" + chr(13) +_
	"<!--" + chr(13) +_
	"theWindow=window.open('','subwin','');" + chr(13) +_
	"theWindow.close();" + chr(13) +_
	"//-->" + chr(13) +_
	"</script>")
end if%>