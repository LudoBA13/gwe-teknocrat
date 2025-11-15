/**
 * Runs when the spreadsheet is opened, creating a custom menu.
 */
function onOpen()
{
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('Teknocrat')
		.addItem('Start', 'showSidebar')
		.addToUi();
}

/**
 * Displays the sidebar for the document generation tool.
 */
function showSidebar()
{
	const html = HtmlService.createHtmlOutputFromFile('sidebar').setTitle('Teknocrat');
	SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Gets or creates a subfolder within a given base folder.
 * @param {GoogleAppsScript.Drive.Folder} baseFolder The base folder to start from.
 * @param {string} path The path to the subfolder, e.g., "Clients/ClientA/Projects".
 * @return {GoogleAppsScript.Drive.Folder} The requested or newly created subfolder.
 */
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

/**
 * Generates multiple Google Docs based on a template and data from a spreadsheet range.
 * @param {string} templateUrl The URL of the template Google Doc.
 * @param {string} dataRange The A1 notation of the data range in the active sheet (e.g., "A1:C10").
 * @param {string} outputFolderId The ID of the Google Drive folder where generated documents will be saved.
 * @param {string} docPath The template for the document name and subfolder path, e.g., "Invoices/<<ClientName>>/Invoice_<<InvoiceNumber>>".
 */
function generateDocuments(templateUrl, dataRange, outputFolderId, docPath)
{
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	const data = sheet.getRange(dataRange).getDisplayValues();
	const headers = data[0];

	const docId        = DocumentApp.openByUrl(templateUrl).getId();
	const srcDoc       = DriveApp.getFileById(docId);

	const outputFolder = DriveApp.getFolderById(outputFolderId);

	for (let i = 1; i < data.length; ++i)
	{
		const rowData = data[i];
		const cancelled = generateDocument(srcDoc, outputFolder, docPath, rowData, headers);
		if (cancelled)
		{
			break;
		}
	}
}

/**
 * Generates a single Google Doc from a template, replacing placeholders with row data.
 * @param {GoogleAppsScript.Drive.File} srcDoc The source template document file.
 * @param {GoogleAppsScript.Drive.Folder} outputFolder The base output folder for generated documents.
 * @param {string} docPath The template for the document name and subfolder path.
 * @param {Array<string>} rowData An array of data for the current row.
 * @param {Array<string>} headers An array of header names corresponding to the rowData.
 * @return {boolean} True if the generation was cancelled, false otherwise.
 */
function generateDocument(srcDoc, outputFolder, docPath, rowData, headers)
{
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

	const tmpName    = 'teknocrat-' + Math.random().toString(36).substring(2);
	const trgDoc     = srcDoc.makeCopy(tmpName);
	const trgDocFile = DocumentApp.openById(trgDoc.getId());

	replacePlaceholdersInDocument(trgDocFile, map);
	trgDocFile.saveAndClose();

	// Check if cancellation was requested
	if (PropertiesService.getUserProperties().getProperty('cancelRequested') === 'true')
	{
		trgDoc.setTrashed(true); // Move to trash
		PropertiesService.getUserProperties().deleteProperty('cancelRequested'); // Clear the flag
		return true; // Signal cancellation
	}

	// Move and rename
	trgDoc.moveTo(trgFolder).setName(docName);
	return false; // Not cancelled
}

/**
 * Replaces placeholders in a Google Document's body with values from a map.
 * Placeholders are expected to be in the format `<<KEY>>`.
 * @param {GoogleAppsScript.Document.Document} doc The document in which to replace placeholders.
 * @param {Map<string, string>} map A map where keys are placeholder names (without `<<>>`) and values are their replacements.
 */
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

/**
 * Replaces placeholders in a string with values from a map.
 * Placeholders are expected to be in the format `<<KEY>>`.
 * @param {string} str The input string containing placeholders.
 * @param {Map<string, string>} map A map where keys are placeholder names (without `<<>>`) and values are their replacements.
 * @return {string} The string with placeholders replaced.
 */
function replacePlaceholdersInString(str, map)
{
	return str.replace(
		/<<([^>]+)>>/g,
		function (match, key)
		{
			return map.get(key) || '';
		}
	);
}

/**
 * Saves user properties for the document generation tool.
 * @param {string} url The template URL.
 * @param {string} range The data range.
 * @param {string} outputFolderId The output folder ID.
 * @param {string} docPath The document path template.
 */
function saveUserProperties(url, range, outputFolderId, docPath)
{
	PropertiesService.getUserProperties().setProperties({
		'teknocrat.templateUrl': url,
		'teknocrat.dataRange': range,
		'teknocrat.outputFolderId': outputFolderId,
		'teknocrat.docPath': docPath
	});
}

/**
 * Retrieves all user properties for the document generation tool.
 * @return {Object<string, string>} An object containing all user properties.
 */
function getUserProperties()
{
	return PropertiesService.getUserProperties().getProperties();
}

/**
 * Sets a flag to indicate that document generation should be cancelled.
 */
function setCancelFlag()
{
	PropertiesService.getUserProperties().setProperty('cancelRequested', 'true');
}