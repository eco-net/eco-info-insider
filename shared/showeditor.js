function showeditor(fldname,fldnum){
	var theurl = '/shared/editor_win.asp?fldname='+fldname
	if(fldnum.length) theurl+='&fldnum='+fldnum
	window.open(theurl,'editor','toolbar=no,location=no,status=yes,menubar=no,scrollbars=nu,resizable=yes,width=500,height=250,top=200,left=100')
}
function showarteditor(fldname,fldnum){
	var theurl = '/shared/editor_win.asp?fldname='+fldname
	if(fldnum.length) theurl+='&fldnum='+fldnum
	window.open(theurl,'editor','toolbar=no,location=no,status=yes,menubar=no,scrollbars=nu,resizable=yes,width=500,height=450,top=200,left=100')
}