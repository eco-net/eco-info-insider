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
	if(theform.title.value.length==0 || theform.author.value.length==0 || theform.shortdescription.value.length==0 ||  theform.description.value.length==0) alert('Alle felter markeret med en * skal udfyldes.')
	else if(elementExsists(theform,'selCat')==true && theform.selCat.selectedIndex==0)alert('Der skal v�lges kategori.')
	else if(elementExsists(theform,'selSbj')==true && theform.selSbj.options.length==0)alert('Der skal v�lges emneord.')
	else if(theform.shortdescription.value.length>200)alert('Den korte beskrivelse er for lang.\nFjern ' + (theform.shortdescription.value.length-250).toString() + ' tegn.');
	else if(theform.keywords.value.length>200)alert('Der er for mange stikord.\nFjern ' + (theform.shortdescription.value.length-250).toString() + ' tegn.');
	else {
		document.validation=true
		if(elementExsists(theform,'selSbj')==true)selectselected(theform.selSbj)
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
