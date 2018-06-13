function openColumns(tablename) {
	window.name='main'
	thewindow=window.open('/dbAdmin/pages/struct_columns.asp?tablename='+tablename,'subwin','scrollbars=yes,resizable=yes,width=800,height=600');
}

function showData(tablename) {
	window.name='main'
	thewindow=window.open('/dbAdmin/pages/cont_data.asp?tablename='+tablename,'subwin','scrollbars=yes,resizable=yes,width=800,height=600');
}


function showInsert(tablename) {
	window.name='main'
	thewindow=window.open('/dbAdmin/pages/struct_getinsert.asp?tablename='+tablename,'subwin','scrollbars=yes,resizable=yes,width=800,height=600');
}
