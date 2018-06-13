<script language="JavaScript">

  var errorString = "Sorry but this web page needs\nWindows95 and Internet Explorer 5 or above to view.";
  var Ok = "false";
  var name =  navigator.appName;
  var version =  parseFloat(navigator.appVersion);
  var platform = navigator.platform;

	if (platform == "Win32" && name == "Microsoft Internet Explorer" && version >= 4){
		Ok = "true";
	} else {
		Ok= "false";
	}

	if (Ok == "false") {
		alert(errorString);
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

	if(document.all.btnSwapView.value != "Html")
	{alert("Rensning for wordkode og <tags> kan kun ske i Normal Visning");}
	else
	{ document.frames("myEditor").document.frames("textEdit").document.body.innerHTML = unescape(s);}
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

	sUrl = document.all.what.value + document.all.url.value;
	doFormat('CreateLink',sUrl);
}

function copyValue() {

	var theHtml = "" + document.frames("myEditor").document.frames("textEdit").document.body.innerHTML + "";
	document.all.EditorValue.value = theHtml;
}

function SwapView_OnClick(){

  if(document.all.btnSwapView.value == "Html"){
		var sMes = "Normal";
		var cMes = "--->";
		var cMes2 = "------->";
    var sStatusBarMes = " Html visning";
	} else {
		var sMes = "Html";
		var cMes = "Rens Word";
		var cMes2 = "Rens pånær linieskift";
    var sStatusBarMes = " Normal visning. Klik Rens for fjernelse af Word kode og <tags>";
  }
	
	document.all.btnSwapView.value = sMes;
	document.all.clean.value = cMes;
	document.all.clean2.value = cMes2;
  window.status  = sStatusBarMes;
	swapMode();
}

function Help_OnClick(){
  window.open("/shared/editor_images/help_document.htm","wHelp", "toolbar=0, scrollbars=yes, width=640, height=480");
}

function OnFormSubmit(){

  //if(confirm("This Document is about to be submitted\nAre you sure you have finished editing?")){
    copyValue();
    //document.fHtmlEditor.submit();
  
}
</script>
<table border="1" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width="50%" height="50%" bordercolor="#CCCCCC" align="center">
<tr valign="top"> 
<td bgcolor="#CCCCCC" width="50%"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>&nbsp;&nbsp;&nbsp;Tekst 
editor. L&aelig;s lige <a href="javascript:Help_OnClick();">hj&aelig;lp</a></b></font></td>
<td bgcolor="#CCCCCC">
<div align="right"> 
<input type="button" name="clean" value="Rens Word" onClick="cleanWord('0');" style="width:80px; font: 8pt verdana;">
<input type="button" name="clean2" value="Rens p&aring;n&aelig;r linieskift" onClick="cleanWord('1');" style="width:130px; font: 8pt verdana;">
</div>
</td>
</tr>
<tr valign="top"> 
<td colspan="2"> 
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
<tr valign="top"> 
<td valign="top"> 
<div id=editbar > 
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="left">
<tr> 
<td> 
<table border="0" cellpadding="0" cellspacing="0" width="271">
<tr> 
<td> 
<table border="0">
<tr valign="baseline"> 
<td nowrap> <img class='clsCursor' src="/shared/editor_images/Cut.gif" width="16" height="16" border="0" alt="Klip " onClick="doFormat('Cut')">&nbsp 
<img class='clsCursor' src="/shared/editor_images/Copy.gif" width="16" height="16" border="0" alt="Kopier" onClick="doFormat('Copy')">&nbsp 
<img class='clsCursor' src="/shared/editor_images/Paste.gif" border="0" alt="Sæt ind" onClick="doFormat('Paste')" width="16" height="16">&nbsp 
</td>
</tr>
</table>
</td>
<td> 
<table border="0">
<tr valign="baseline"> 
<td nowrap> <img class='clsCursor' src="/shared/editor_images/para_bul.gif" width="16" height="16" border="0" alt="Punkttegn" onClick="doFormat('InsertUnorderedList');" >&nbsp 
<img class='clsCursor' src="/shared/editor_images/para_num.gif" width="16" height="16" border="0" alt="Nummereret liste" onClick="doFormat('InsertOrderedList');" >&nbsp 
<img class='clsCursor' src="/shared/editor_images/indent.gif" width="20" height="16" alt="Indryk" onClick="doFormat('Indent')">&nbsp 
<img class='clsCursor' src="/shared/editor_images/outdent.gif" width="20" height="16" alt="Ryk ud" onClick="doFormat('Outdent')">&nbsp 
<img class='clsCursor' src="/shared/editor_images/hr.gif" width="16" height="18" alt="Vandret streg" onClick="doFormat('InsertHorizontalRule')">&nbsp 
</td>
</tr>
</table>
</td>
<td> 
<table border="0">
<tr> 
<td nowrap valign="baseline"> 
<div align="left"> Font: 
<select name="size" onChange="doFormat('FontSize',document.all.size.value);" style="font: 8pt verdana;">
<option value="None" selected>Str</option>
<option value="1">1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
<option value="6">6</option>
<option value="7">7</option>
<option value="+1">+1</option>
<option value="+2">+2</option>
<option value="+3">+3</option>
<option value="+4">+4</option>
<option value="+5">+5</option>
<option value="+6">+6</option>
<option value="+7">+7</option>
</select>
<img class='clsCursor' src="/shared/editor_images/Bold.gif" width="16" height="16" border="0" align="absmiddle" alt="Fed" onClick="doFormat('Bold')">&nbsp 
<img class='clsCursor' src="/shared/editor_images/Italics.gif" width="16" height="16" border="0" align="absmiddle" alt="Kursiv" onClick="doFormat('Italic')">&nbsp 
<img class='clsCursor' src="/shared/editor_images/underline.gif" width="16" height="16" border="0" align="absmiddle" alt="Understreget" onClick="doFormat('Underline')" >&nbsp 
<img class='clsCursor' src="/shared/editor_images/left.gif" width="16" height="16" border="0" alt="Venstrestillet" align="absmiddle"  onClick="doFormat('JustifyLeft')"> 
<img class='clsCursor' src="/shared/editor_images/centre.gif" width="16" height="16" border="0" alt="Centreret" align="absmiddle" onClick="doFormat('JustifyCenter')">&nbsp 
<img class='clsCursor' src="/shared/editor_images/right.gif" width="16" height="16" border="0" alt="Højrestillet" align="absmiddle"  onClick="doFormat('JustifyRight')">&nbsp 
</div>
</td>
</tr>
</table>
</td>
</tr>
</table>
</td>
<td align="left" nowrap valign="top"> 
<input type="button" name="btnSwapView" value="Html" onClick="SwapView_OnClick();" style="width:50px; font: 8pt verdana;">
</td>
</tr>
</table>
</div>
</td>
</tr>
<tr valign="top" align="left"> 
<td valign="top"> 
<table width="100%" border="0" height="100%">
<tr valign="top"> 
<td width="100%" height="100%"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="400">
<tr valign="top"> 
<td bgcolor="#FFFFFF"><iframe id=myEditor src="/shared/pd_edit.htm" onFocus="initToolBar(this)" width=100% height=100%></iframe></td>
</tr>
</table>
</td>
<!-- COLORPICKER_GOES_HERE -->
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>


<script>
  initToolBar("foo");
  window.status  = " Normal visning. Klik Rens for fjernelse af Word kode og <tags>";
</script>
