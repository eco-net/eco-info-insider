function record(theform) {
	document.validation=false
	if(theform.title.value.length==0 || theform.startdate.value.length==0 || theform.shortdescription.value.length==0 ||  theform.description.value.length==0 || theform.fulladress.value.length==0 || theform.postnr.value.length==0 || theform.selCat.options.length==0 || theform.selSbj.options.length==0) alert('Alle felter markeret med en * skal udfyldes.')
	else if(theform.shortdescription.value.length>200)alert('Den korte beskrivelse er for lang.\nFjern ' + (theform.shortdescription.value.length-250).toString() + ' tegn.');
	else if(theform.keywords.value.length>200)alert('Der er for mange stikord.\nFjern ' + (theform.shortdescription.value.length-250).toString() + ' tegn.');
	else {
		document.validation=true
		selectSelected(theform.selSbj)
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
