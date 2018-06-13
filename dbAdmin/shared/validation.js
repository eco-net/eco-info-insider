function feedback() {
	thewindow=window.open('/feedback/new.asp','fbck','scrollbars=yes,resizable=yes,width=500,height=350');
}

function validateSelect(theObj,objName) {
	var isValid
	var theValue=theObj.options[theObj.selectedIndex].value
	if (theValue==0 || theValue=='') {
		isValid=false
		document.MM_errortext +='- ' + objName + ' skal vælges.\n'
	} else {valid=true }
	document.MM_returnValue = isValid
}

function evalValidation() {
	if(!document.MM_returnValue) alert(document.MM_errortext)
}

function YY_checkform() { //v3.05
  var args = YY_checkform.arguments; var myDot=true; myV=''; var myErr='';var addErr=false;
  if (document.all){eval("args[0]=args[0].replace(/.layers/gi, '.all');");}
  for (var i=1; i<args.length;i=i+4){
    if (args[i+1].charAt(0)=='#'){
      var myReq=true; args[i+1]=args[i+1].substring(1);
    }else{myReq=false}
    var myObj = eval(args[0]+'.'+args[i])
    if (myObj.type=='text'){
      if (myReq&&myObj.value.length==0){addErr=true}
      myV=myObj.value;
      if ((myV.length>0)&&(args[i+2]==1)){ //fromto
        if (isNaN(parseInt(myV))||myV<args[i+1].substring(0,args[i+1].indexOf('_'))/1||myV > args[i+1].substring(args[i+1].indexOf('_')+1)/1){addErr=true}
      }
      if ((myV.length>0)&&(args[i+2]==2)){ //e-mail
        if (myV.lastIndexOf('.')<myV.lastIndexOf('@')||myV.lastIndexOf('.')==-1||myV.lastIndexOf('@')==-1){addErr=true}
      }
      if ((myV.length>0)&&(args[i+2]==3)){ // date
        var myD=''; myM=''; myY=''; myYY=0; myDot=true;
        for(var j=0;j<args[i+1].length;j++){
          if(args[i+1].charAt(j)=='D')myD=myD.concat(myObj.value.charAt(j));
          if(args[i+1].charAt(j)=='M')myM=myM.concat(myObj.value.charAt(j));
          if(args[i+1].charAt(j)=='Y'){myY=myY.concat(myObj.value.charAt(j)); myYY++}
          if(args[i+1].charAt(j)=='-'&&myObj.value.charAt(j)!='-')myDot=false;
          if(args[i+1].charAt(j)=='.'&&myObj.value.charAt(j)!='.')myDot=false;
          if(args[i+1].charAt(j)=='/'&&myObj.value.charAt(j)!='/')myDot=false;
        }
        if(myD/1<1||myD/1>31||myM/1<1||myM/1>12||myY.length!=myYY)myDot=false;
        if(!myDot){addErr=true}
       }
      if ((myV.length>0)&&(args[i+2]==4)){ // time
        var myDot=true;
        var myH = myObj.value.substr(0,myObj.value.indexOf(':'))/1;
        var myM = myObj.value.substr(myObj.value.indexOf(':')+1,2)/1;
                var myP = myObj.value.substr(myObj.value.indexOf(':')+3,2);
        if ((args[i+1])=="12:00pm"){if(myH<0||myH>12||myM<0||myM>59||(myP!="pm"&&myP!="am")||myObj.value.length>7)myDot=false; }
        if ((args[i+1])=="12:00"){if(myH<0||myH>12||myM<0||myM>59||myObj.value.length>5)myDot=false;}
        if ((args[i+1])=="24:00"){if(myH<0||myH>23||myM<0||myM>59||myObj.value.length>5)myDot=false;}
        if(!myDot){addErr=true}
      }
      if ((myV.length>0)&&(args[i+2]==5)){ // check this 2
        if (!eval(args[0]+'.'+args[i+1]+'.checked')){addErr=true}
      }
    }
    if (myObj.type=='radio'){
      if (args[i+2]==1&&myObj.checked&&eval(args[0]+'.'+args[i+1]+'.value.length')/1==0){addErr=true}
      if (args[i+2]==2){
        myDot=false;
        myV=eval(args[0]+'.'+args[i].substring(0,args[i].lastIndexOf('[')));
        for(var j=0;j<myV.length;j++){myDot=myDot||myV[j].checked}
        if(!myDot){myErr+='* ' +args[i+3]+'\n'}
      }
    }
    if (myObj.type=='checkbox'){
      if(args[i+2]==1&&myObj.checked==false){addErr=true}
      if(args[i+2]==2&&myObj.checked&&eval(args[0]+'.'+args[i+1]+'.value.length')/1==0){addErr=true}
      if (args[i+2]==3){
        myDot=false;
        myV=eval(args[0]+'.'+args[i].substring(0,args[i].lastIndexOf('[')));
        for(var j=0;j<myV.length;j++){myDot=myDot||myV[j].checked}
        if(!myDot){myErr+='* ' +args[i+3]+'\n'}
      }
    }
    if (myObj.type=='select-one'||myObj.type=='select-multiple'){
      if(args[i+2]==1&&eval(args[0]+'.'+args[i]+'.selectedIndex')/1==0){addErr=true}
    }
    if (myObj.type=='textarea'){
      myV = eval(args[0]+'.'+args[i]+'.value');
      if(myV.length<args[i+1]){addErr=true}
    }
    if (addErr){myErr+='* '+args[i+3]+'\n'; addErr=false}
  }
  if (myErr!=''){alert('Der er følgende fejl i indtastningen:\t\t\t\t\t\n\n'+myErr)}
  document.MM_returnValue = (myErr=='');
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}