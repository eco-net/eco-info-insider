function toNumber(thestr) {
	var theval=''
	for(var i=0;i<thestr.length;i++){
		if(!isNaN(thestr.charAt(i)))theval+=thestr.charAt(i)
	}
	return theval
}

function elementExsists(theform,elementname,indexNum) {
	var bool=false
	if(theform.elements[elementname]!=null) {
		if(!indexNum)bool=true;
		else {if(theform.elements[elementname].length>=indexNum)bool=true;else bool=false;}
	}
	return bool
}

function record(theform) {
	document.validation=false
	if(theform.firstname.value.length==0 || theform.lastname.value.length==0 || theform.shortdescription.value.length==0 ||  theform.description.value.length==0 || theform.street1.value.length==0 || theform.zip.value.length==0|| theform.emailaddress1.value.length==0) alert('Alle felter markeret med en * skal udfyldes.')
	else if(elementExsists(theform,'selCat')==true && theform.selCat.options.length==0)alert('Der skal v�lges kategorier.')
	else if(elementExsists(theform,'selSbj')==true && theform.selSbj.options.length==0)alert('Der skal v�lges emneord.')
	else if(theform.shortdescription.value.length>200)alert('Den korte beskrivelse er for lang.\nFjern ' + (theform.shortdescription.value.length-250).toString() + ' tegn.');
	else if(theform.keywords.value.length>200)alert('Der er for mange stikord.\nFjern ' + (theform.shortdescription.value.length-250).toString() + ' tegn.');
	else if(theform.keywords.value.length>200)alert('Der er for mange stikord.\nFjern ' + (theform.shortdescription.value.length-250).toString() + ' tegn.');
	else if(elementExsists(theform,'username')==true && (theform.username.value.length>20 || theform.username.value.length<4))alert('Brugernavn skal v�re p� mellem 4 og 20 tegn')
	else if(elementExsists(theform,'password')==true && (theform.password.value.length>8 || theform.password.value.length<4))alert('Adgangskode skal v�re p� mellem 4 og 8 tegn')
	else {
		document.validation=true
		if(elementExsists(theform,'selCat')==true)selectselected(theform.selCat)
		if(elementExsists(theform,'selSbj')==true)selectselected(theform.selSbj)
		theform.zip_dk.value=toNumber(theform.zip.value)
	}
}

function donew(theform) {
	record(theform)
	if(document.validation==true) {
		theform.action='act_new.asp';
	}
}

function doedit(theform) {
	record(theform)
	if(document.validation==true) {
		theform.action='act_edit.asp';
	}
}
