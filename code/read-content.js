function createFile(exportTable) {
    var tab_text = '<html xmlns:x="urn:schemas-microsoft-com:office:excel">';
    tab_text = tab_text + '<head><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>';
    tab_text = tab_text + '<x:Name>Sheet1</x:Name>';
    tab_text = tab_text + '<x:WorksheetOptions><x:Panes></x:Panes></x:WorksheetOptions></x:ExcelWorksheet>';
    tab_text = tab_text + '</x:ExcelWorksheets></x:ExcelWorkbook></xml>';
	tab_text = tab_text + '<style>td {mso-number-format:"\@";/*force text*/}</style></head>';
    tab_text = tab_text + "<body><table border='1px'>";
    tab_text = tab_text + exportTable.outerHTML;
    tab_text = tab_text + '</table></body>';
    tab_text = tab_text + '</html>';
    
    //Save the file
	var excelBlob = new Blob([tab_text], { type: "application/vnd.ms-excel;charset=utf-8" });
    var link=window.URL.createObjectURL(excelBlob);
	window.location =link;
	window.URL.revokeObjectURL(link);
}

function exportToExcel(){
	const selection = document.getSelection();
	let containerTagName = '';

	if (selection.rangeCount === 0) {
		alert("no selection found");
		return false;
	}

	const selectionRange = selection.getRangeAt(0); // Only consider the first range

	const container = selectionRange.commonAncestorContainer;

	if(container instanceof Element && (container.tagName.toLowerCase() == "tbody" || container.tagName.toLowerCase() == "table")){

		createFile(container);
	}
	else
	{
		alert("table not selected");
	}
}

exportToExcel();