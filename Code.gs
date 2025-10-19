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

	for (let i = 1; i < data.length; i++)
	{
		const rowData = data[i];

		const docId      = DocumentApp.openByUrl(templateUrl).getId();
		const newDoc     = DriveApp.getFileById(docId).makeCopy();
		const newDocFile = DocumentApp.openById(newDoc.getId());
		const body       = newDocFile.getBody();

		for (let j = 0; j < headers.length; j++)
		{
			const placeholder = '<<' + headers[j] + '>>';
			body.replaceText(placeholder, rowData[j]);
		}

		newDocFile.setName('Generated Doc ' + i);
		newDocFile.saveAndClose();
	}
}