function GP_AdvOpenWindow(theURL,winName,features,popWidth,popHeight,winAlign,ignorelink,alwaysOnTop,autoCloseTime,borderless) { //v2.0
	var leftPos=0,topPos=0,autoCloseTimeoutHandle, ontopIntervalHandle, w = 480, h = 340;  
	if (popWidth > 0) features += (features.length > 0 ? ',' : '') + 'width=' + popWidth;
	if (popHeight > 0) features += (features.length > 0 ? ',' : '') + 'height=' + popHeight;

	if (winAlign && winAlign != "" && popWidth > 0 && popHeight > 0) {
		if (document.all || document.layers || document.getElementById) {w = screen.availWidth; h = screen.availHeight;}
		if (winAlign.indexOf("center") != -1) {topPos = (h-popHeight)/2;leftPos = (w-popWidth)/2;}
		if (winAlign.indexOf("bottom") != -1) topPos = h-popHeight; if (winAlign.indexOf("right") != -1) leftPos = w-popWidth; 
		if (winAlign.indexOf("left") != -1) leftPos = 0; if (winAlign.indexOf("top") != -1) topPos = 0; 						
		features += (features.length > 0 ? ',' : '') + 'top=' + topPos+',left='+leftPos;
	}

	if (document.all && borderless && borderless != "" && features.indexOf("fullscreen") != -1) features+=",fullscreen=1";
	if (window["popupWindow"] == null) window["popupWindow"] = new Array();
	var wp = popupWindow.length;
	popupWindow[wp] = window.open(theURL,winName,features);
	if (popupWindow[wp].opener == null) popupWindow[wp].opener = self;  
	if (document.all || document.layers || document.getElementById) {
		if (borderless && borderless != "") {popupWindow[wp].resizeTo(popWidth,popHeight); popupWindow[wp].moveTo(leftPos, topPos);}
		if (alwaysOnTop && alwaysOnTop != "") {
			ontopIntervalHandle = popupWindow[wp].setInterval("window.focus();", 50);
			popupWindow[wp].document.body.onload = function() {window.setInterval("window.focus();", 50);}; 
		}

		if (autoCloseTime && autoCloseTime > 0) {
			popupWindow[wp].document.body.onbeforeunload = function() {
				if (autoCloseTimeoutHandle) window.clearInterval(autoCloseTimeoutHandle);
				window.onbeforeunload = null;
			}
			autoCloseTimeoutHandle = window.setTimeout("popupWindow["+wp+"].close()", autoCloseTime * 1000);
		}
		window.onbeforeunload = function() {for (var i=0;i<popupWindow.length;i++) popupWindow[i].close();};
	}
	document.MM_returnValue = (ignorelink && ignorelink != "") ? false : true;
}

function showpageurl(){
	window.open('/general/sel_link.asp?urlfld=url','Link','width=300,height=200,top=200,left=200')
}


function ColorPalette_OnClick(colorString){
	cpick.bgColor=colorString;
	document.all.colourp.value=colorString;
	doFormat('ForeColor',colorString);
}

function initToolBar(ed) {
	var eb = document.all.editbar;
	if (ed!=null) {
		eb._editor = window.frames['myEditor'];
	}
}

function doFormat(what) {
	var eb = document.all.editbar;
	if(what == "FontName"){
		if(arguments[1] != 1){
			eb._editor.execCommand(what, arguments[1]);
			document.all.font.selectedIndex = 0;
		} 
	} else if(what == "FontSize"){
		if(arguments[1] != 1){
		  eb._editor.execCommand(what, arguments[1]);
		  document.all.size.selectedIndex = 0;
		} 
	} else {
	   eb._editor.execCommand(what, arguments[1]);
	}
}

function swapMode() {
	var eb = document.all.editbar._editor;
  eb.swapModes();
}

function create() {
    var eb = document.all.editbar;
    eb._editor.newDocument();
}

function newFile(){
	create();
}

function makeUrl(){
	if(document.all.url.value.substr(0,1)=='/')sUrl = document.all.url.value;
	else sUrl = document.all.what.value + document.all.url.value;
	doFormat('CreateLink',sUrl);
}

function cleanuptext() {
	var theText=document.frames("myEditor").document.frames("textEdit").document.body.innerHTML
//	while (theText.indexOf('<p>')>-1) {
//		theText=theText.replace('<p>','#p#')
//	}
//	while (theText.indexOf('<P>')>-1) {
//		theText=theText.replace('<P>','#p#')
//	}
//	while (theText.indexOf('</p>')>-1) {
//		theText=theText.replace('</p>','#/p#')
//	}
//	while (theText.indexOf('</P>')>-1) {
//		theText=theText.replace('</P>','#/p#')
//	}
	while (theText.indexOf('<br>')>-1) {
		theText=theText.replace('<br>','#br#')
	}
	while (theText.indexOf('<Br>')>-1) {
		theText=theText.replace('<Br>','#br#')
	}
	while (theText.indexOf('<BR>')>-1) {
		theText=theText.replace('<BR>','#br#')
	}
	document.frames("myEditor").document.frames("textEdit").document.body.innerHTML=theText;
	document.frames("myEditor").document.frames("textEdit").document.body.innerHTML = document.frames("myEditor").document.frames("textEdit").document.body.innerText
	theText=document.frames("myEditor").document.frames("textEdit").document.body.innerText;
	while (theText.indexOf('#p#')>-1) {
		theText=theText.replace('#p#','<p>')
	}
	while (theText.indexOf('#br#')>-1) {
		theText=theText.replace('#br#','<br>')
	}
	while (theText.indexOf('#/p#')>-1) {
		theText=theText.replace('#/p#','</p>')
	}
	document.frames("myEditor").document.frames("textEdit").document.body.innerHTML=theText
}

function copyValue() {
	if(document.htmlview)swapMode();
	var theHtml = "" + document.frames("myEditor").document.frames("textEdit").document.body.innerHTML + "";
	theHtml=setTargetOnLinks(theHtml);
	document.all.bodyText[document.thenum].value = theHtml;
}

function setTargetOnLinks(thetext){
	var i,tStart, y
	var fStr, eStr, etext
	etext = thetext.toLowerCase();
	tStart = etext.indexOf('href="http:\/\/')
	while (tStart>0){
		for(i=tStart+13;i<thetext.length;i++){
			if(thetext.charAt(i)=='"'){
				fStr = etext.substr(i+1,etext.indexOf('>',i+1)-i);
				if(fStr.indexOf('target=')==-1){
					fStr = thetext.substr(0,i+1)
					eStr = thetext.substr(i+1,thetext.length-1)
					thetext=fStr+' target="new"'+eStr
				}
				y=i
				i=thetext.length+1
			}
		}
		etext=thetext.toLowerCase();
		tStart = etext.indexOf('href="http:\/\/',y+1)
	}
	return thetext
}


function SwapView_OnClick(){

  if(document.all.btnSwapView.value == "HTML-visning"){
		var sMes = "Designvisning";
		document.htmlview=true
	    var sStatusBarMes = "Viser nu html-koder";
	} else {
		var sMes = "HTML-visning"
	    var sStatusBarMes = "Viser nu designvisning";
		document.htmlview=false
  }
	document.all.btnSwapView.value = sMes;
	window.status  = sStatusBarMes;
	swapMode();
}

function Help_OnClick(){
  window.open("/includes/editor/help_document.htm","wHelp", "toolbar=0, scrollbars=yes, width=640, height=480");
}

function OnFormSubmit(){
  if(confirm("Du er ved at sende dokumentet.\nBekræft venligst at du er færdig med at redigere")){
    copyValue();
    document.fHtmlEditor.submit();
  }
}
function cleanWord(newline)
{
var str = "" + document.frames("myEditor").document.frames("textEdit").document.body.innerHTML + "";
	
var s,start,slut,read
s=""
read=1
start="<"
slut=">"
P="<P>"
//marker <P> og <BR>
if (newline=='1') 
{

while (str.indexOf('<P')>-1) {
str=str.replace("<P","#P#<");}
while (str.indexOf('</P>')>-1) {
str=str.replace("</P>","#/P#");}
while (str.indexOf('<BR>')>-1) {
str=str.replace("<BR>","#BR#");}
}
	for (i=0;i<str.length;i++)
	{
		if (str.charAt(i)==start){read=0;}//fra < til > læses ikke
		if (read==1){s=s + str.charAt(i)}// der læses
		if (str.charAt(i)==slut){read=1;}//herefter læses der igen
	}
//ret mrk til tag
if (newline=='1') 
{	
while (s.indexOf('#P#')>-1) {
s=s.replace("#P#","<P>");}
while (s.indexOf('#/P#')>-1) {
s=s.replace("#/P#","</P>");}
while (s.indexOf('#BR#')>-1) {
s=s.replace("#BR#","<BR>");}
}

	if(document.all.btnSwapView.value != "HTML-visning")
	{alert("Rensning for wordkode og <tags> kan kun ske i Normal Visning");}
	else
	{ document.frames("myEditor").document.frames("textEdit").document.body.innerHTML = unescape(s);}
}
