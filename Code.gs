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

function getOrCreateSubFolder(baseFolder, path)
{
	const parts = path.split('/');
	let currentFolder = baseFolder;
	for (const part of parts)
	{
		// Ignore empty parts from trailing slashes etc.
		if (part)
		{
			const folders = currentFolder.getFoldersByName(part);
			if (folders.hasNext())
			{
				currentFolder = folders.next();
			}
			else
			{
				currentFolder = currentFolder.createFolder(part);
			}
		}
	}
	return currentFolder;
}

function generateDocuments(templateUrl, dataRange, outputFolderId, docPath)
{
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	const data = sheet.getRange(dataRange).getValues();
	const headers = data[0];

	const outputFolder = DriveApp.getFolderById(outputFolderId);

	for (let i = 1; i < data.length; ++i)
	{
		const rowData = data[i];
		const map = new Map;
		for (let j = 0; j < headers.length; ++j)
		{
			map.set(headers[j], rowData[j]);
		}

		// Resolve placeholders in the doc path
		let resolvedDocPath = docPath;
		for (const [key, value] of map.entries())
		{
			resolvedDocPath = resolvedDocPath.replace(new RegExp('<<' + key + '>>', 'g'), value);
		}

		const pathParts     = resolvedDocPath.split('/');
		const docName       = pathParts.pop();
		const subfolderPath = pathParts.join('/');
		const trgFolder     = getOrCreateSubFolder(outputFolder, subfolderPath);

		const docId      = DocumentApp.openByUrl(templateUrl).getId();
		const tmpName    = 'teknocrat-' + Math.random().toString(36).substring(2);
		const tmpDoc     = DriveApp.getFileById(docId).makeCopy(tmpName);
		const trgDocFile = DocumentApp.openById(tmpDoc.getId());
		
		replacePlaceholdersInDocument(trgDocFile, map);
		trgDocFile.saveAndClose();

		// Move and rename
		tmpDoc.moveTo(trgFolder).setName(docName);
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

function saveUserProperties(url, range, outputFolderId, docPath)
{
	PropertiesService.getUserProperties().setProperties({
		'teknocrat.templateUrl': url,
		'teknocrat.dataRange': range,
		'teknocrat.outputFolderId': outputFolderId,
		'teknocrat.docPath': docPath
	});
}

function getUserProperties()
{
	return PropertiesService.getUserProperties().getProperties();
}