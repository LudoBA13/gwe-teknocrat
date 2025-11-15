/**
 * A class for generating and manipulating Google Documents.
 */
class DocumentGenerator
{
	/**
	 * @private
	 * @type {GoogleAppsScript.Drive.File}
	 */
	templateFile;

	/**
	 * @private
	 * @type {GoogleAppsScript.Document.Document}
	 */
	templateDocument;

	/**
	 * @private
	 * @type {string}
	 */
	placeholderStart;

	/**
	 * @private
	 * @type {string}
	 */
	placeholderEnd;

	/**
	 * @private
	 * @type {RegExp}
	 */
	placeholderRegex;

	/**
	 * @private
	 * @type {string}
	 */
	escapedPlaceholderStart;

	/**
	 * @private
	 * @type {string}
	 */
	escapedPlaceholderEnd;

	/**
	 * @private
	 * @type {Set<string>}
	 */
	placeholders;

	/**
	 * @param {string} docUrlOrId The URL or ID of the Google Document.
	 * @param {string=} placeholderStart Optional. The starting delimiter for placeholders. Defaults to '<<'.
	 * @param {string=} placeholderEnd Optional. The ending delimiter for placeholders. Defaults to '>>'.
	 */
	constructor(docUrlOrId, placeholderStart = '<<', placeholderEnd = '>>')
	{
		if (docUrlOrId.startsWith('https://'))
		{
			const doc = DocumentApp.openByUrl(docUrlOrId);
			this.templateFile = DriveApp.getFileById(doc.getId());
		}
		else
		{
			this.templateFile = DriveApp.getFileById(docUrlOrId);
		}

		this.templateDocument = DocumentApp.openById(this.templateFile.getId());

		this.placeholderStart = placeholderStart;
		this.placeholderEnd = placeholderEnd;

		this.escapedPlaceholderStart = this._escapeRegExp(placeholderStart);
		this.escapedPlaceholderEnd = this._escapeRegExp(placeholderEnd);

		this.placeholderRegex = new RegExp(this.escapedPlaceholderStart + '(.*?)' + this.escapedPlaceholderEnd, 'g');

		this.placeholders = this._getPlaceholders();
	}

	/**
	 * Escapes special characters in a string for use in a regular expression.
	 * @param {string} str The string to escape.
	 * @return {string} The escaped string.
	 * @private
	 */
	_escapeRegExp(str)
	{
		return str.replace(/[.*+?^${}()|[\\]/g, '\\$&'); // $& means the whole matched string
	}

	/**
	 * Returns the currently managed template file.
	 * @return {GoogleAppsScript.Drive.File} The Google Drive File.
	 */
	getTemplateFile()
	{
		return this.templateFile;
	}

	/**
	 * Returns the ID of the currently managed template file.
	 * @return {string} The ID of the Google Drive File.
	 */
	getTemplateFileId()
	{
		return this.templateFile.getId();
	}

	/**
	 * Replaces all occurrences of a specific placeholder key with a given value in the managed document.
	 * @param {string} key The placeholder key (e.g., "name" for "<<name>>").
	 * @param {string} value The value to replace the placeholder with.
	 */
	replacePlaceholder(key, value)
	{
		const body = this.templateDocument.getBody();
		const specificPlaceholderRegex = new RegExp(
			this.escapedPlaceholderStart +
			this._escapeRegExp(key) +
			this.escapedPlaceholderEnd,
			'g'
		);
		body.replaceText(specificPlaceholderRegex, value);
	}

	/**
	 * Collects all unique placeholder keys from the document.
	 * @return {Set<string>} A set of unique placeholder keys.
	 * @private
	 */
	_getPlaceholders()
	{
		const text = this.templateDocument.getBody().getText();
		const matches = text.matchAll(this.placeholderRegex);
		return new Set(Array.from(matches, match => match[1]));
	}

	/**
	 * Generates a new document from the template, replacing placeholders with the provided variables.
	 * @param {Iterable<[string, string]>} vars An iterable of key-value pairs for placeholder replacement.
	 * @return {GoogleAppsScript.Document.Document} The newly generated document.
	 */
	generateDocument(vars)
	{
		// 1. Create a copy
		const templateFile = DriveApp.getFileById(this.templateFile.getId());
		const newFile = templateFile.makeCopy();
		const outputDocument = DocumentApp.openById(newFile.getId());
		const body = outputDocument.getBody();

		// 2. Iterate and replace from vars
		for (const [key, value] of vars)
		{
			const specificPlaceholderRegex = new RegExp(
				this.escapedPlaceholderStart +
				this._escapeRegExp(key) +
				this.escapedPlaceholderEnd,
				'g'
			);
			body.replaceText(specificPlaceholderRegex, value);
		}

		// 3. Replace remaining placeholders with empty string
		body.replaceText(this.placeholderRegex, '');

		// 4. Return the new document
		return outputDocument;
	}

	/**
	 * Saves and closes the currently managed document.
	 */
	saveAndCloseDocument()
	{
		this.templateDocument.saveAndClose();
	}
}
