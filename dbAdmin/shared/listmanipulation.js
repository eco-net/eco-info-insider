function WA_AddValueToList(ListObj,TextString,ValString,Position)  {
  if (isNaN(parseInt(Position)))   {
    Position = ListObj.options.length;
  }
  else  {
    Position = parseInt(Position);
  }
  if (ListObj.length > Position)  {
  ListObj.options[Position].text=TextString;
  if (ValString != "")  {
    ListObj.options[Position].value = ValString;
  }
    else  {
      ListObj.options[Position].value=TextString;
    }
  }
  else  {
    var LastOption = new Option();
    var OptionPosition = ListObj.options.length;
    ListObj.options[OptionPosition] = LastOption;
    ListObj.options[OptionPosition].text = TextString;
    if (ValString != "")  {
      ListObj.options[OptionPosition].value = ValString;
    }
    else  {
      ListObj.options[OptionPosition].value=TextString;
    }
  }
}

function WA_subAwithBinC(a,b,c)
{

	var i = c.indexOf(a);
	var l = b.length;

	while (i != -1)	{
		c = c.substring(0,i) + b + c.substring(i + a.length,c.length);  //replace all valid a values with b values in the selected string c.
  i += l
		i = c.indexOf(a,i);
	}
	return c;

}

function WA_AddSubToSelected(sublist,targetlist,repeatvalues,leavetop,leavebottom,noseltop,noselbot,topval,toptext)     {
  for (var j=0; j<noseltop; j++)     {
    sublist.options[j].selected = false;
  }
  for (var k=0; k<noselbot; k++)     {
    sublist.options[sublist.options.length-(k+1)].selected = false;
  }
  if (sublist.selectedIndex >= 0)      {
   if (leavebottom)     {
      botrec = new Array(2);
      botrec[0] = targetlist.options[targetlist.options.length-1].value;
      botrec[1] = targetlist.options[targetlist.options.length-1].text;
      targetlist.options[targetlist.options.length-1] = null;
    }
    if (!leavetop && targetlist.options.length > 0)     {
      if (targetlist.options[0].value == topval)     {
        targetlist.options[0] = null;  
      }
    }
    else     {
      if (leavetop && toptext != '')     {
        targetlist.options[0].value = topval;
        targetlist.options[0].text = toptext;
      }
    }
    for (var o=0; o<sublist.options.length; o++)     {
      if (sublist.options[o].selected && o >= noseltop && o < sublist.options.length - noselbot)     {
        theText = sublist.options[o].text;
        theValue = sublist.options[o].value;
        addvalue = true;
        if (!repeatvalues)      {
          for (var p=0; p<targetlist.options.length; p++)     {
            if (theValue == targetlist.options[p].value)      {
              addvalue = false;
            }
          }
        }
        if (addvalue)  WA_AddValueToList(targetlist,theText,theValue,targetlist.options.length);
      }
    }
    if (leavebottom)     {
      WA_AddValueToList(targetlist,botrec[1],botrec[0],targetlist.options.length);
    }
  }
  for (var l=0; l < targetlist.options.length; l++)    {
    targetlist.options[l].selected = false;
  }
}

function WA_RemoveSelectedFromList(theBox,nottoremove,noneselectedoption,noneselectedvalue,noneselectedtext)     {
  var n=0;
  var selectedArray = new Array();
  for (var j=0; j<theBox.options.length; j++)     {
    if (!theBox.options[j].selected || nottoremove.indexOf("|WA|" + theBox.options[j].value + "|WA|") >= 0)     {
      theBox.options[n].value = theBox.options[j].value;
      theBox.options[n].text = theBox.options[j].text;
      n++;
    }
    else {
	    selectedArray[selectedArray.length] = j;
    }
  }
  for (var k=0; k<selectedArray.length; k++)  {
    theBox.options[selectedArray[k]].selected = false;
  }
  m = n;
  while (m<=j)     {
    theBox.options[n] = null;
    m++;
  }
  if (theBox.options.length == noneselectedoption && noneselectedtext != "")     {
    noneselectedvalue = WA_subAwithBinC("|WA|",",",noneselectedvalue);
    noneselectedtext = WA_subAwithBinC("|WA|",",",noneselectedtext);
    WA_AddValueToList(theBox,noneselectedtext,noneselectedvalue,0);
  }
  for (var l=0; l < theBox.options.length; l++)    {
    theBox.options[l].selected = false;
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