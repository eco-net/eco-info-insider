function elementExsists(theform,elementname,indexNum) {
	var bool=false
	if(theform.elements[elementname]!=null) {
		if(!indexNum)bool=true;
		else {if(theform.elements[elementname].length>=indexNum)bool=true;else bool=false;}
	}
	return bool
}
function doedit(theform) {
	if(elementExsists(theform,'selSbj')==true)
	{selectselected(theform.selSbj)}
	 {
		theform.action='act_art_kap.asp';
	}
}

