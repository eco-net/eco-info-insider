//////////////////////////////////////////////////////////////////////////////////////
//GENERAL UTILITIES

function checkImg(theField,theHeight,theWidth) {
	var error=0;
	var re = new RegExp(".(gif|jpg|png|bmp|jpeg)$","i");
	if(!re.test(theField.value) && theField.value.length>0){error=1;alert('Det valgte billede kan ikke vises i en browser.')}
	if(error==0){
		var imgURL = 'file:///' + theField.value.replace(/\\/gi,'/');
		theField.gp_img = new Image();
		theField.gp_img.src=imgURL;
		if(theField.gp_img.width!=theWidth && theWidth>0) error=1;
		if(theField.gp_img.height!=theHeight && theHeight>0) error=1;
		if(error==1)alert('Det valgte billede har ikke de korrekte dimensioner. (' + theField.gp_img.width.toString() + ',' + theField.gp_img.height.toString() + ')');
	};
	if(error==0)return true
	else {return false}
}


function validate_choose(theform){
	document.validation=true;
	var spl1=theform.ba1.options[theform.ba1.selectedIndex].value
	if(spl1==0){alert('Vælg en Branding Area.');document.validation=false}
}

function dochoose(theform){
	validate_choose(theform);
	if(document.validation==true){theform.action='act_choose.asp'}
}

function record(theform,isedit) {
	document.validation=false
	var imgOK=false
	if(isedit==1 && theform.file.value.length==0)imgOK=true;
	if(imgOK==false && theform.file.value.length==0)alert('Der skal vælges et billede.');
	if(theform.file.value.length>0)imgOK=checkImg(theform.file,200,258);
	if(imgOK==true){
		if(theform.headline.value.length==0 || theform.body.value.length==0) alert('Både overskrift og tekst skal udfyldes.')
//		else if(theform.link.value.length==0) alert('Der skal vælges en link.')
		else {document.validation=true}
	}
}

function donew(theform) {
	record(theform,0)
	if(document.validation==true) {
		theform.action='act_new.asp';
	}
}

function doedit(theform) {
	record(theform,1)
	if(document.validation==true) {
		theform.action='act_edit.asp';
	}
}
