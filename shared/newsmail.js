function newsMail(what,mail,site) {
	if(mail.indexOf('@')==-1 || mail.indexOf('.')==-1) {
		alert('Indtast en gyldig email-adresse !');
		document.newsmail.themail.select		
	}
	else {
		window.name="main"
		thewindow=window.open('/home/newsmail.asp?what='+what+'&mail='+mail+'&site='+site,'subwin','scrollbars=no,resizable=no,width=400,height=150,top=200,left=200');
		document.newsmail.themail.value=''
	}
	return false
}
