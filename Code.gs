function onOpen()
{
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('Teknocrat')
		.addItem('Start', 'showSidebar')
		.addToUi();
}

function showSidebar()
{
	const html = HtmlService.createHtmlOutputFromFile('sidebar').setTitle('Teknocrat');
	SpreadsheetApp.getUi().showSidebar(html);
}

function generateDocuments(templateUrl, dataRange)
{
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	const data = sheet.getRange(dataRange).getValues();

	const headers = data[0];

	for (let i = 1; i < data.length; ++i)
	{
		const rowData = data[i];
		const map = new Map();
		for (let j = 0; j < headers.length; ++j)
		{
			map.set(headers[j], rowData[j]);
		}

		const docId      = DocumentApp.openByUrl(templateUrl).getId();
		const newDoc     = DriveApp.getFileById(docId).makeCopy();
		const newDocFile = DocumentApp.openById(newDoc.getId());
		
		replacePlaceholdersInDocument(newDocFile, map);

		newDocFile.setName('Generated Doc ' + i);
		newDocFile.saveAndClose();
	}
}

function replacePlaceholdersInDocument(doc, map)
{
	const body = doc.getBody();
	const replacements = [];
	const searchPattern = '<<.*?>>';

	// 1. Find all placeholders and store them with their replacement values
	let searchResult = body.findText(searchPattern);
	while (searchResult)
	{
		const rangeElement = searchResult.getElement();
		const text = rangeElement.asText().getText();
		const placeholder = text.substring(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive() + 1);
		const key = placeholder.substring(2, placeholder.length - 2);
		const value = map.get(key) || '';

		replacements.push({
			rangeElement: rangeElement,
			start: searchResult.getStartOffset(),
			end: searchResult.getEndOffsetInclusive(),
			value: value
		});

		searchResult = body.findText(searchPattern, searchResult);
	}

	// 2. Perform the replacements in reverse order
	for (let i = replacements.length - 1; i >= 0; --i)
	{
		const r = replacements[i];
		const textElement = r.rangeElement.asText();
		textElement.deleteText(r.start, r.end);
		textElement.insertText(r.start, r.value);
	}
}