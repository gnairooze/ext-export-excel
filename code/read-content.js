function createHtml(exportTable) {
    let tab_text = `
	<html xmlns:x="urn:schemas-microsoft-com:office:excel">
	<head>
		<xml>
			<x:ExcelWorkbook>
				<x:ExcelWorksheets>
					<x:ExcelWorksheet>
						<x:Name>Sheet1</x:Name>
						<x:WorksheetOptions><x:Panes></x:Panes></x:WorksheetOptions>
					</x:ExcelWorksheet>
				</x:ExcelWorksheets>
			</x:ExcelWorkbook>
		</xml>
		<style>td {mso-number-format:"\@";/*force text*/}</style>
	</head>
	<body>
		<table border='1px'>
			{{EXPORT_TABLE}}
		</table>
	</body>
	</html>
	`;
	
	return tab_text.replace("{{EXPORT_TABLE}}", exportTable.outerHTML);
}

function saveFile(content){
	var excelBlob = new Blob([content], { type: "application/vnd.ms-excel;charset=utf-8" });
    var link=window.URL.createObjectURL(excelBlob);
	window.location =link;
	window.URL.revokeObjectURL(link);
}

function exportToExcel(){
	const selection = document.getSelection();

	if (selection.rangeCount === 0) {
		alert("no selection found");
		return false;
	}

	const container = selection.getRangeAt(0).commonAncestorContainer;

	if(container instanceof Element && (container.tagName.toLowerCase() == "tbody" || container.tagName.toLowerCase() == "table")){
		let content = createHtml(container);
		saveFile(content);
	}
	else
	{
		alert("table not selected");
	}
}

exportToExcel();