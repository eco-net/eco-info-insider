function openColumns(tablename) {
	window.name='main'
	thewindow=window.open('struct_columns.asp?tablename='+tablename,'subwin','scrollbars=yes,resizable=yes,width=800,height=600');
}

function showData(tablename) {
	window.name='main'
	thewindow=window.open('cont_data.asp?tablename='+tablename,'subwin','scrollbars=yes,resizable=yes,width=800,height=600');
}

function dropTable(tablename) {
	if(confirm('Do you really want to delete the table "' + tablename + '" ?')) {
		window.location.href='struct_tables.asp?tablename='+tablename+'&drop=1'
	}
}

function emptyTable(tablename) {
	if(confirm('Do you really want to delete all ata from the table "' + tablename + '" ?')) {
		window.location.href='struct_tables.asp?tablename='+tablename+'&empty=1'
	}
}

function showInsert(tablename) {
	window.name='main'
	thewindow=window.open('struct_getinsert.asp?tablename='+tablename,'subwin','scrollbars=yes,resizable=yes,width=800,height=600');
}
