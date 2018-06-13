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
	if(theform.title.value.length==0 || theform.startdate.value.length==0 ||  theform.description.value.length==0 || theform.postnr.value.length==0) alert('Alle felter markeret med en * skal udfyldes.')
	else if(elementExsists(theform,'selCat')==true && theform.selCat.options.length==0)alert('Der skal vælges kategori.')
	else if(elementExsists(theform,'selSbj')==true && theform.selSbj.options.length==0)alert('Der skal vælges emneord.')
	else if(theform.shortdescription.value.length>250)alert('Den korte beskrivelse er for lang.\nFjern ' + (theform.shortdescription.value.length-250).toString() + ' tegn.');
	else if(theform.keywords.value.length>200)alert('Der er for mange stikord.\nFjern ' + (theform.keywords.value.length-200).toString() + ' tegn.');
	else {
		document.validation=true
		if(elementExsists(theform,'selSbj')==true)selectselected(theform.selSbj)
		if(elementExsists(theform,'selCat')==true)selectselected(theform.selCat)
	}
}

function donew(theform) {

	record(theform)
	if(document.validation==true) {
		if(!theform.enddate.value.length)theform.enddate.value=theform.startdate.value
		theform.action='act_new.asp';
	}
}

function doedit()
{
	theform=mf;
	record(theform)
	if(document.validation==true){
		collectOrgid_edit(theform);
		theform.action='act_edit.asp';
		//theform.submit();
	}
}

function collectOrgid_edit(theform) 
{
	if(elementExsists(theform,"theorgid2"))
	{
		for(var i=0;i<theform.theorgid2.length;i++){
			theform.theorgid.value+=','+theform.theorgid2[i].value
		}
		theform.theorgid.value=theform.theorgid.value.substring(1);
	}
}
