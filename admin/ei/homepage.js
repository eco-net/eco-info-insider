//////////////////////////////////////////////////////////////////////////////////////
//GENERAL UTILITIES

function checkImg(theField,theHeight,theWidth) {
	var error=0;
	var re = new RegExp(".(gif|jpg|png|bmp|jpeg)$","i");
	if(!re.test(theField.value) && theField.value.length>0){error=1;alert('The chosen image can not be displayed in a browser.')}
	if(error==0){
		var imgURL = 'file:///' + theField.value.replace(/\\/gi,'/');
		theField.gp_img = new Image();
		theField.gp_img.src=imgURL;
		if(theField.gp_img.width!=theWidth && theWidth>0) error=1;
		if(theField.gp_img.height!=theHeight && theHeight>0) error=1;
		if(error==1)alert('The chosen image is not of the correct dimensions.');
	};
	if(error==0)return true
	else {return false}
}

//////////////////////////////////////////////////////////////////////////////////////
//HOMEPAGE

function choose_spl() {
	this.location='spl/choose.asp'
}

function new_spl() {
	this.location='spl/new.asp'
}

function edit_spl(theForm) {
	var theVal=theForm.sel_edit_spl.options[theForm.sel_edit_spl.selectedIndex].value
	if(theVal>0) this.location='spl/edit.asp?id='+theVal
	else alert('Please choose a splash in the drop down.')
}

function delete_spl(theForm) {
	var theVal=theForm.del_spl.options[theForm.del_spl.selectedIndex].value
	if(theVal>0) this.location='spl/delete.asp?id='+theVal
	else  alert('Please choose a Splash in the drop down.')
}

function choose_ba() {
	this.location='ba/choose.asp'
}

function new_ba() {
	this.location='ba/new.asp'
}

function edit_ba(theForm) {
	var theVal=theForm.sel_edit_ba.options[theForm.sel_edit_ba.selectedIndex].value
	if(theVal>0) this.location='ba/edit.asp?id='+theVal
	else alert('Please choose a Branding Area in the drop down.')
}

function delete_ba(theForm) {
	var theVal=theForm.del_ba.options[theForm.del_ba.selectedIndex].value
	if(theVal>0) this.location='ba/delete.asp?id='+theVal
	else  alert('Please choose a Branding Area in the drop down.')
}

function choose_news() {
	this.location='choosenews.asp'
}

//////////////////////////////////////////////////////////////////////////////////////
// HOMEPAGE CONTENT FORMS

function validate_choosespl(theform){
	document.validation=true;
	var spl1=theform.spl1.options[theform.spl1.selectedIndex].value
	var spl2=theform.spl2.options[theform.spl2.selectedIndex].value
	var spl3=theform.spl3.options[theform.spl3.selectedIndex].value
	if((spl1==0 && spl2==0) || (spl1==0 && spl3==0) || (spl2==0 && spl3==0)){alert('Please choose at least 2 splashes.');document.validation=false}
	else if(spl1==spl2 || spl1==spl3 || spl2==spl3) {alert('Please choose different splashes.');document.validation=false}
}

function choosespl(theform){
	validate_choosespl(theform);
	if(document.validation==true){theform.action='act_choose.asp';theform.submit()}
}

function chooseba(theform) {
	document.validation=true;
	if(theform.sel.options.length==0){alert('Please select at least one Branding Area.');document.validation=false}
	if(document.validation==true){selectselected(theform.sel);theform.action='act_choose.asp';theform.submit()}
}

function validate_choosenews(theform) {
	document.validation=true;
	var tbg=theform.tbgid.options[theform.tbgid.selectedIndex].value
	if(theform.method[1].checked==true && tbg==0) {alert('Please choose 1 - 3 News to display.');document.validation=false}
	if(theform.method[0].checked==true)theform.tbgid.selectedIndex=0;
}

function choosenews(theform) {
	validate_choosenews(theform);
	if(document.validation==true){theform.action='act_choosenews.asp';theform.submit()}
}

function previewHomepage(theform,what){
	//what=1 - featured links
	//what=2 - splashes
	//what=3 - topbargraphic
	if(what==1)validate_choosefl(theform);
	else if(what==2)validate_choosespl(theform);
	else if(what==3)validate_choosetbg(theform)
	if(document.validation==true) {
		//var thewin=window.open('../../previewhp/getparams.asp?pt='+what,'Preview','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=yes')
		var thewin=window.open('/shared/getparams.asp?pt='+what,'Preview','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=yes')
	}
}

function previewHomepage_setparams(newform,what){
	var theform=opener.document.forms[0]
	if(what==1){
		newform.fl1.value=theform.fl1.options[theform.fl1.selectedIndex].value;
		newform.fl2.value=theform.fl2.options[theform.fl2.selectedIndex].value;
		newform.fl3.value=theform.fl3.options[theform.fl3.selectedIndex].value;
	} else if(what==2){
		newform.spl1.value=theform.spl1.options[theform.spl1.selectedIndex].value;
		newform.spl2.value=theform.spl2.options[theform.spl2.selectedIndex].value;
		newform.spl3.value=theform.spl3.options[theform.spl3.selectedIndex].value;
	} else if(what==3){
		if(theform.method[0].checked==true)newform.method.value=1;
		else{
			newform.method.value=0;
			newform.tbgid.value=theform.tbgid.options[theform.tbgid.selectedIndex].value
		}
	}
	newform.submit();
}

//////////////////////////////////////////////////////////////////////////////////////
// SPLASHES

function splrecord(theform,isedit) {
	document.validation=false
	var theerr=0
	var imgOK=checkImg(theform.file,80,129);
	if(imgOK==true){
		if(theform.headline.value.length==0 || theform.content.value.length==0) alert('Please fill in both headline and content.')
		else if(theform.pagelink.value.length==0) alert('Please choose a page to link to.')
		else if(theform.pagelink.value.indexOf('areaerror')>-1) alert('The selected Area has no page.\nPlease select a SubArea.')
		else {document.validation=true}
	}
}

function newspl(theform) {
	splrecord(theform,0)
	if(document.validation==true) {
		theform.action='act_new.asp';theform.encoding='multipart/form-data';theform.submit();
	}
}

function editspl(theform) {
	splrecord(theform,1)
	if(document.validation==true) {
		theform.action='act_edit.asp';theform.encoding='multipart/form-data';theform.submit();
	}
}

//////////////////////////////////////////////////////////////////////////////////////
// BRANDING AREAS

function barecord(theform,isedit) {
	document.validation=false
	var theerr=0
	var imgOK=checkImg(theform.file,200,583);
	if(imgOK==true){
		if(theform.headline.value.length==0 || theform.content.value.length==0) alert('Please fill in both headline and content.')
		else if(theform.pagelink.value.length==0) alert('Please choose a page to link to.')
		else if(theform.pagelink.value.indexOf('areaerror')>-1) alert('The selected Area has no page.\nPlease select a SubArea.')
		else {document.validation=true}
	}
}

function newba(theform) {
	barecord(theform,0)
	if(document.validation==true) {
		theform.action='act_new.asp';theform.encoding='multipart/form-data';theform.submit();
	}
}

function editba(theform) {
	barecord(theform,1)
	if(document.validation==true) {
		theform.action='act_edit.asp';theform.encoding='multipart/form-data';theform.submit();
	}
}
