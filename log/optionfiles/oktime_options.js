document.write('<option value="0" selected>&lt;V&aelig;lg&gt;</option>');
document.write('<option value="10">Kommende arrangementer</option>');
document.write('<option value="20">Afholdte arrangementer</option>');

idag=new Date();
m=idag.getMonth();
y=idag.getYear();
if(y<1900){y=y+1900;}//Firefox

for(var i=0;i<12;i++){
writeMonth(m,y);
m++;
if (m>11){
m=0;
y++;
}
}

function writeMonth(month,year){
		
	mno=month+1;
	opt=year + "-" + mno + '-01';
	//optionstr=opt.toString();
	optionstr=String(opt)
	val=mname(month) + ' ' + year;
	//valuestr=val.toString();
	valuestr=String(val);
document.write('<option value="' + optionstr + '">' + valuestr + '</option>');

}

function mname(month){

if (month==0){return "Januar"}
if (month==1){return "Februar"}
if (month==2){return "Marts"}
if (month==3){return "April"}
if (month==4){return "Maj"}
if (month==5){return "Juni"}
if (month==6){return "Juli"}
if (month==7){return "August"}
if (month==8){return "September"}
if (month==9){return "Oktober"}
if (month==10){return "November"}
if (month==11){return "December"}
}
