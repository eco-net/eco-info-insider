var array1081=  new Array("('<Vælg>','0',true,true)","('Aalborg Kommune','851')","('Brønderslev-Dronninglund Kommune','810')","('Frederikshavn Kommune','813')","('Hjørring Kommune','860')","('Jammerbugt Kommune','849')","('Læsø Kommune','825')","('Mariagerfjord Kommune','846')","('Morsø Kommune','773')","('Rebild Kommune','840')","('Thisted Kommune','787')","('Vesthimmerland Kommune','820')")
var array1082=  new Array("('<Vælg>','0',true,true)","('Århus kommune','751')","('Favrskov Kommune','710')","('Hedensted Kommune','766')","('Herning Kommune','657')","('Holstebro Kommune','661')","('Horsens Kommune','615')","('Ikast-Brande Kommune','756')","('Lemvig Kommune','665')","('Norddjurs Kommune','707')","('Odder Kommune','727')","('Randers Kommune','730')","('Ringkøbing-Skjern Kommune','760')","('Samsø Kommune','741')","('Silkeborg Kommune','740')","('Skanderborg Kommune','746')","('Skive kommune','779')","('Struer kommune','671')","('Syddjurs Kommune','706')","('Viborg Kommune','791')")
var array1083=  new Array("('<Vælg>','0',true,true)","('Aabenraa Kommune','580')","('Ærø Kommune','492')","('Assens Kommune','420')","('Billund Kommune','530')","('Esbjerg Kommune','561')","('Faaborg-Midtfyn Kommune','430')","('Fanø Kommune','563')","('Fredericia Kommune','607')","('Haderslev Kommune','510')","('Kerteminde Kommune','440')","('Kolding Kommune','621')","('Langeland Kommune','482')","('Middelfart Kommune','410')","('Nordfyns Kommune','480')","('Nyborg Kommune','450')","('Odense Kommune','461')","('Sønderborg Kommune','540')","('Svendborg Kommune','479')","('Tønder Kommune','550')","('Varde Kommune','573')","('Vejen Kommune','575')","('Vejle Kommune','630')")
var array1084=  new Array("('<Vælg>','0',true,true)","('Albertslund Kommune','165')","('Allerød Kommune','201')","('Ballerup Kommune','151')","('Bornholm Kommune','400')","('Brøndby Kommune','153')","('Dragør Kommune','155')","('Egedal Kommune','240')","('Fredensborg Kommune','210')","('Frederiksberg Kommune','147')","('Frederikssund Kommune','250')","('Frederiksværk-Hundested Kommune','260')","('Furesø Kommune','190')","('Gentofte Kommune','157')","('Gladsaxe Kommune','159')","('Glostrup Kommune','161')","('Gribskov Kommune','270')","('Helsingør Kommune','217')","('Herlev Kommune','163')","('Hillerød Kommune','219')","('Høje-Taastrup','169')","('Hørsholm Kommune','223')","('Hvidovre Kommune','167')","('Ishøj Kommune','183')","('Københavns Kommune','101')","('Lyngby-Taarbæk Kommune','173')","('Rødovre Kommune','175')","('Rudersdal Kommune','230')","('Tårnby Kommune','185')","('Vallensbæk Kommune','187')")
var array1085=  new Array("('<Vælg>','0',true,true)","('Faxe Kommune','320')","('Greve Kommune','253')","('Guldborgsund Kommune','376')","('Holbæk Kommune','316')","('Kalundborg Kommune','326')","('Køge Kommune','259')","('Lejre Kommune','350')","('Lolland Kommune','360')","('Næstved Kommune','370')","('Odsherred Kommune','306')","('Ringsted Kommune','329')","('Roskilde Kommune','265')","('Slagelse Kommune','330')","('Solrød Kommune','269')","('Sorø Kommune','340')","('Stevns Kommune','336')","('Vordingborg kommune','390')")
function populateSel(theSel,selected) {
	if(selected>0){
		var selectedArray = eval("array"+selected);
		while (selectedArray.length < theSel.options.length) {
			theSel.options[(theSel.options.length - 1)] = null;
		}
		for (var i=0; i < selectedArray.length; i++) {
			theSel.options[i]=eval("new Option" + selectedArray[i]);
		}
	} else {
		while (theSel.options.length>0) {
			theSel.options[(theSel.options.length - 1)] = null;
		}
		theSel.options[0]=eval("new Option" + "('<--','0',true,true)");
	}
}