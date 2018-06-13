function doSearch(theform) { //execute the relevant pages when generel search are used
	var theval=theform.thesection.options[theform.thesection.selectedIndex].value
	if (theval==1) //searching all apps
		{theform.action='/home/searchall.asp'}
	else if(theval==2) //searching de grønne sider
		{theform.thetext.name='dgstext';theform.action='/dgs/list.asp'}
	else if(theval==3) // searching øko-kalenderen
		{theform.thetext.name='oktext';theform.action='/ok/list.asp'}
	else if(theval==4) // searching Det grønne bibliotek 
		{theform.thetext.name='dgbtext';theform.action='/dgb/list.asp'}
	else if(theval==5) // searching opslagstavlen
		{theform.action='/opsl/list.asp'}
	else if(theval==6) // searching nyheder
		{theform.thetext.name='';theform.action='/news/list.asp'}
	return true
}

function mailcolleague(){
 window.open('/home/mailpage.asp','mailpage','width=300,height=220,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,copyhistory=no')
 }

function sendemail(theemail){
	window.opener.location.href="mailto:"+theemail+theblah
	this.close()
}

function bookmark(){
 window.external.AddFavorite(location.href,document.title)
 }

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function checkPostnr(val) { //v2.0
var zip=val
var length=zip.length
var abc=zip.substr(4)
//der er skrevet by
if(abc.length>0){alert("Skriv kun postnr")}
else{
var nr=parseInt(zip)
//første tal er et bogstav
	if (isNaN(nr)){	alert("Skriv postnr")}
else{
//der er skrevet mindre end 4 tal
	if(length<4){alert("Skriv 4 tal")}
}}}
