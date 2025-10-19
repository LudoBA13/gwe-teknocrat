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
		const resolvedDocPath = replacePlaceholdersInString(docPath, map);

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

		// Check if cancellation was requested
		if (PropertiesService.getUserProperties().getProperty('cancelRequested') === 'true')
		{
			tmpDoc.setTrashed(true);
			PropertiesService.getUserProperties().deleteProperty('cancelRequested');
			break;
		}

		// Move and rename
		tmpDoc.moveTo(trgFolder).setName(docName);
	}
}

function replacePlaceholdersInDocument(doc, map)
{
	const body           = doc.getBody();
	const replacementMap = new Map;
	for (const m of body.getText().matchAll(/<<([^>]+)>>/g))
	{
		const key     = m[1];
		const match   = m[0].replace(/[.\\+*?^$()\[\]{}|]/g, '\\$&');
		const replace = map.get(key) || '';

		replacementMap.set(match, replace);
	}

	for (const [match, replace] of replacementMap.entries())
	{
		body.replaceText(match, replace);
	}
}

function replacePlaceholdersInString(str, map)
{
	let result = str;
	for (const [key, value] of map.entries())
	{
		result = result.replace(new RegExp('<<' + key + '>>', 'g'), value);
	}
	return result;
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

function setCancelFlag()
{
	PropertiesService.getUserProperties().setProperty('cancelRequested', 'true');
}