function addValue(fromlist,tolist){
	var Position = tolist.options.length;
	var LastOption = new Option();
	tolist.options[Position] = LastOption;
	tolist.options[Position].text = fromlist.options[fromlist.selectedIndex].text;
	tolist.options[Position].value = fromlist.options[fromlist.selectedIndex].value;
}

function removeValue(fromlist){
	var n=0
	var selectedArray = new Array();
	for (var j=0; j<fromlist.options.length; j++)     {
		if (!fromlist.options[j].selected){
			fromlist.options[n].value = fromlist.options[j].value;
			fromlist.options[n].text = fromlist.options[j].text;
			n++;
		} else {
			selectedArray[selectedArray.length] = j;
		}
	}
	for (var k=0; k<selectedArray.length; k++)  {
		fromlist.options[selectedArray[k]].selected = false;
	}
	m = n;
	while (m<=j)     {
		fromlist.options[n] = null;
		m++;
	}
}

function WA_MoveSelectedInList(sourceselect,tomove,topnum,botnum)     {
  var selectedIndex = sourceselect.selectedIndex;
  if (selectedIndex > topnum && tomove == "0" && selectedIndex <= sourceselect.options.length-botnum-1)    {
    oldvals = new Array(sourceselect.options[selectedIndex - 1].value, sourceselect.options[selectedIndex - 1].text);
    sourceselect.options[selectedIndex-1].value = sourceselect.options[selectedIndex].value;
    sourceselect.options[selectedIndex-1].text  = sourceselect.options[selectedIndex].text;
    sourceselect.options[selectedIndex].value   = oldvals[0];
	sourceselect.options[selectedIndex].text    = oldvals[1];
    sourceselect.selectedIndex                  = selectedIndex-1;
  }
  if (selectedIndex < sourceselect.options.length-botnum-1 && tomove == "1" && selectedIndex >= topnum)     {
    oldvals = new Array(sourceselect.options[selectedIndex + 1].value, sourceselect.options[selectedIndex + 1].text);
    sourceselect.options[selectedIndex+1].value = sourceselect.options[selectedIndex].value;
    sourceselect.options[selectedIndex+1].text  = sourceselect.options[selectedIndex].text;
    sourceselect.options[selectedIndex].value   = oldvals[0];
	sourceselect.options[selectedIndex].text    = oldvals[1];
    sourceselect.selectedIndex                  = selectedIndex+1;
  }
}

function selectselected(thelist) {
	var i
	for(i=0; i < thelist.options.length; i++) {
		thelist.options[i].selected=true
	}
}