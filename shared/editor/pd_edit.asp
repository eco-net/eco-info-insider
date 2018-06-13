<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<style>
  BODY {margin-left: 0px;margin-top: 0px; margin: 0px;padding: 0px; border: none}
  IFRAME {width: 100%; height: 100%; margin-left: 0px;margin-top: 0px; padding: 0px}
</style>
<script>
// Default format is WYSIWYG HTML
var format="HTML"

// Set the focus to the editor
function setFocus() {
 textEdit.focus()
}

// Execute a command against the editor
// At minimum one argument is required. Some commands
// require a second optional argument:
// eg., ("formatblock","<H1>") to make an H1

function execCommand(command) {
textEdit.focus()
 if (format=="HTML") {
  var edit = textEdit.document.selection.createRange()
  if (arguments[1]==null)
   edit.execCommand(command)
  else
   edit.execCommand(command,false, arguments[1])
  edit.select()
  textEdit.focus()
 }
}

// Selects all the text in the editor
function selectAllText(){

var edit = textEdit.document;

edit.execCommand('SelectAll');
textEdit.focus();

}

function newDocument() {
	textEdit.document.open()
	textEdit.document.write("")
	textEdit.document.close()
	textEdit.focus()
}

function loadDoc(htmlString) {
	textEdit.document.open()
	textEdit.document.write(htmlString)
	textEdit.document.close()
}
// Initialize the editor with an empty document
function initEditor() {
	var htmlString = parent.document.all.bodyText[<%=request("thenum")%>].value;
	textEdit.document.designMode="On"
	textEdit.document.open()
	textEdit.document.write(htmlString)
	textEdit.document.close()
	swapModes()
	swapModes()
	textEdit.document.createStyleSheet("/shared/styles.css")
	textEdit.focus()
}

// Swap between WYSIWYG mode and raw HTML mode
function swapModes() {
	if (format=="HTML") {
		textEdit.document.body.innerText = textEdit.document.body.innerHTML
		format="Text"
 	}
	else {
   		textEdit.document.body.innerHTML = textEdit.document.body.innerText
   		format="HTML"
 	}
	// textEdit.focus()
	var s = textEdit.document.body.createTextRange()
	s.collapse(false)
	s.select()
}

function cleanuptext() {alert("clean");
	textEdit.document.body.innerHTML = textEdit.document.body.innerText
}

window.onload = initEditor


</script>
<!-- Turn off the outer body scrollbars. The IFRAME editbox will display its 
     own scrollbars when necessary -->
<body scroll="No" class="nomarg">

<iframe id="textEdit" MARGINHEIGHT="2" MARGINWIDTH="2" border="0" frameborder="0">
</iframe>
<script>
textEdit.focus();
</script>
</body>
</html>
